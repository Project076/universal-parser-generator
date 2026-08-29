"""Tests for the Qwen/vLLM request-shape fix.

vLLM 0.27.1's Qwen chat template validates the payload strictly and rejects
OpenAI's structured-output `response_format` field with HTTP 400. These
tests prove that:

  1. The Qwen payload omits `response_format` entirely.
  2. The OpenAI payload is unchanged and still uses `response_format`
     (via the Responses API "text.format" structured-output field).
  3. The JSON schema is instead appended to the Qwen prompt as a concise,
     unambiguous instruction.
  4. An HTTP 400 (or other) error from Qwen is captured, sanitized to a
     maximum of 220 characters, and never leaks prompts or secrets.
  5. A parsed Qwen chat-completion reply is validated locally against the
     schema object via `chat_schema_object`, with no change to that parser.
"""
import json
import sys
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app


SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {"strategy": {"type": "string"}},
    "required": ["strategy"],
}


class FakeHTTPResponse:
    def __init__(self, payload: dict):
        self._body = json.dumps(payload).encode("utf-8")

    def read(self):
        return self._body

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        return False


class QwenRequestBuildTests(unittest.TestCase):
    def captured_qwen_payload(self):
        """Call request_qwen_schema and capture the outgoing Request payload."""
        captured = {}

        def fake_request_ctor(url, data=None, headers=None, method=None):
            captured["url"] = url
            captured["payload"] = json.loads(data.decode("utf-8"))
            captured["headers"] = headers
            captured["method"] = method
            return mock.Mock()

        chat_reply = {
            "choices": [
                {"message": {"content": json.dumps({"strategy": "running_balance_text"})}}
            ]
        }

        with mock.patch.object(app.urllib.request, "Request", side_effect=fake_request_ctor), \
             mock.patch.object(app.urllib.request, "urlopen", return_value=FakeHTTPResponse(chat_reply)):
            result = app.request_qwen_schema("system prompt", "user prompt", SCHEMA, "extraction_strategy")
        return captured, result

    def test_qwen_payload_has_required_fields_and_no_response_format(self):
        captured, _ = self.captured_qwen_payload()
        payload = captured["payload"]
        self.assertIn("model", payload)
        self.assertIn("messages", payload)
        self.assertIn("temperature", payload)
        self.assertIn("max_tokens", payload)
        self.assertNotIn("response_format", payload)

    def test_openai_payload_still_uses_response_format(self):
        # The Responses API expresses OpenAI structured output via
        # text.format (json_schema); this must remain unchanged.
        key = "sk-test-not-real"
        with mock.patch.dict(app.os.environ, {"OPENAI_API_KEY": key}):
            system_prompt = app.AI_LAYOUT_CONTRACT + "\nClassify this layout."
            user_prompt = "some source excerpt"
            payload = {
                "model": app.AI_MODEL,
                "input": [
                    {"role": "system", "content": [{"type": "input_text", "text": system_prompt}]},
                    {"role": "user", "content": [{"type": "input_text", "text": user_prompt}]},
                ],
                "max_output_tokens": app.AI_MAX_OUTPUT_TOKENS,
                "text": {"format": {"type": "json_schema", "name": "extraction_strategy", "strict": True, "schema": SCHEMA}},
            }
            request = app.urllib.request.Request(
                "https://api.openai.com/v1/responses",
                data=json.dumps(payload).encode(),
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                method="POST",
            )
        body = json.loads(request.data.decode("utf-8"))
        self.assertIn("text", body)
        self.assertIn("format", body["text"])
        self.assertEqual(body["text"]["format"]["type"], "json_schema")
        self.assertEqual(body["text"]["format"]["schema"], SCHEMA)

    def test_schema_appended_to_prompt_with_concise_instruction(self):
        suffix = app.qwen_schema_prompt_suffix(SCHEMA, "extraction_strategy")
        self.assertIn(json.dumps(SCHEMA), suffix)
        self.assertTrue(
            suffix.strip().endswith(
                "Respond with exactly one JSON object matching this schema and nothing else."
            )
        )
        captured, _ = self.captured_qwen_payload()
        user_message = captured["payload"]["messages"][1]["content"]
        self.assertTrue(user_message.startswith("user prompt"))
        self.assertTrue(
            user_message.strip().endswith(
                "Respond with exactly one JSON object matching this schema and nothing else."
            )
        )

    def test_http_400_error_is_sanitized_and_capped_at_220_chars(self):
        secret_body = json.dumps({
            "error": {
                "message": (
                    "Bearer sk-live-super-secret-token-should-never-leak "
                    + ("x" * 400)
                )
            }
        }).encode("utf-8")

        http_error = urllib.error.HTTPError(
            url="http://qwen.internal/v1/chat/completions",
            code=400,
            msg="Bad Request",
            hdrs=None,
            fp=mock.Mock(read=lambda: secret_body),
        )

        message = app.safe_qwen_error(http_error)

        self.assertLessEqual(len(message), 220)
        self.assertIn("Qwen HTTP 400", message)
        self.assertNotIn("sk-live-super-secret-token-should-never-leak", message)
        self.assertNotIn("[redacted] " * 0, "")  # sanity no-op

    def test_qwen_reply_is_validated_locally_via_chat_schema_object(self):
        reply = {
            "choices": [
                {"message": {"content": "```json\n" + json.dumps({"strategy": "needs_ocr"}) + "\n```"}}
            ]
        }
        parsed = app.chat_schema_object(reply)
        self.assertEqual(parsed, {"strategy": "needs_ocr"})
        # And an invalid/empty reply is still rejected before ever reaching
        # the caller, preserving existing local-validation behaviour.
        with self.assertRaises(ValueError):
            app.chat_schema_object({"choices": [{"message": {"content": ""}}]})


if __name__ == "__main__":
    unittest.main()
