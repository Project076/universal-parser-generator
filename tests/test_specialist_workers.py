import unittest

from specialist_workers import (
    STEP_DOMAINS,
    build_blocked_step_escalation,
    build_specialist_worker_registry,
)


MANIFEST = {number: f"STEP_{number}" for number in range(1, 51)}


class SpecialistWorkerTests(unittest.TestCase):
    def test_every_step_has_a_unique_scoped_worker(self):
        registry = build_specialist_worker_registry(MANIFEST)
        self.assertEqual(set(registry), set(range(1, 51)))
        self.assertEqual(len({worker["worker_id"] for worker in registry.values()}), 50)
        self.assertTrue(all(worker["library_scopes"] for worker in registry.values()))
        self.assertTrue(all(not worker["may_overwrite_certified_learning"] for worker in registry.values()))

    def test_narration_and_furniture_are_separate_specialists(self):
        registry = build_specialist_worker_registry(MANIFEST)
        self.assertEqual(STEP_DOMAINS[7], "narration")
        self.assertEqual(STEP_DOMAINS[8], "furniture")
        self.assertNotEqual(registry[7]["library_scopes"], registry[8]["library_scopes"])

    def test_blocked_step_replays_only_from_its_owner(self):
        worker = build_specialist_worker_registry(MANIFEST)[7]
        event = build_blocked_step_escalation(
            worker,
            "continuation",
            ["continuation_merge", "narration_source_cell"],
            ["narration"],
            ["continuation"],
        )
        self.assertEqual(event["blocked_step"], "S07_STEP_7")
        self.assertEqual(event["replay_from_step"], 7)
        self.assertEqual(event["learning_policy"], "versioned_addendum_only")
        self.assertIn("advance_to_unrelated_step", event["forbidden"])

    def test_incomplete_manifest_fails_closed(self):
        with self.assertRaises(ValueError):
            build_specialist_worker_registry({1: "ONLY_ONE"})


if __name__ == "__main__":
    unittest.main()
