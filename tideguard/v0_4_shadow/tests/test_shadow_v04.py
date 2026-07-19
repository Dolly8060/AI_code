import json
import unittest
from pathlib import Path

from tideguard_shadow.engine import (
    acceleration_label,
    ai_earnings_momentum,
    ai_funding_burden,
    composite_z_state,
    evaluate_shadow_v04,
    funding_cost_state,
    plumbing_liquidity_state,
    pre_reopen_quality_blocked,
    risk_transmission_state,
    slope_label,
)


class ShadowV04Tests(unittest.TestCase):
    def test_plumbing_states(self):
        self.assertEqual(plumbing_liquidity_state(3, 0, -1)["state"], "NORMAL")
        self.assertEqual(plumbing_liquidity_state(7, 0, -1)["state"], "WATCH")
        self.assertEqual(
            plumbing_liquidity_state(7, 2_000_000_000, -1)["state"],
            "STRESS",
        )
        self.assertEqual(
            plumbing_liquidity_state(None, 0, None)["state"],
            "DATA_INSUFFICIENT",
        )

    def test_funding_cost(self):
        self.assertEqual(
            funding_cost_state({"2y": 85, "10y": 84, "30y": 70, "tips10y": 89})[
                "state"
            ],
            "HIGH",
        )
        self.assertEqual(
            funding_cost_state({"2y": 10, "10y": 15, "30y": 20, "tips10y": 40})[
                "state"
            ],
            "LOW",
        )
        self.assertEqual(
            funding_cost_state({"2y": 50, "10y": None})["state"],
            "DATA_INSUFFICIENT",
        )

    def test_composite_z(self):
        self.assertEqual(composite_z_state({"a": 1, "b": 1, "c": 1})["state"], "HIGH")
        self.assertEqual(
            composite_z_state({"a": 2, "b": 2, "c": 2})["state"],
            "EXTREME",
        )
        self.assertEqual(
            composite_z_state({"a": 0, "b": None})["state"],
            "DATA_INSUFFICIENT",
        )

    def test_risk_transmission(self):
        out = risk_transmission_state(
            red_tail=False,
            vix=19,
            hy_oas_5d_change_bp=2,
            hy_oas_2d_change_bp=1,
            ai_semi_state="FAIL",
            repair_pass_count=4,
            a_share_state="COOLING",
            funding_cost="HIGH",
            duration_cost="HIGH",
        )
        self.assertEqual(out["state"], "RATES_TO_EQUITY")

    def test_quality_block(self):
        out = pre_reopen_quality_blocked(
            "PRE_REOPEN_LOW",
            "HIGH",
            True,
            "NORMAL",
            "DATA_INSUFFICIENT",
            False,
        )
        self.assertTrue(out["blocked"])
        self.assertIn("DURATION_COST_HIGH", out["reasons"])
        self.assertIn("SMH_HARD_FAIL", out["reasons"])

    def test_momentum_labels(self):
        self.assertEqual(slope_label(0.4, 0.3), "RISING")
        self.assertEqual(slope_label(-0.4, 0.3), "FALLING")
        self.assertEqual(
            acceleration_label(0.5, 0.1, 0.3),
            "ACCELERATING",
        )

    def test_ai_overlays(self):
        self.assertEqual(
            ai_earnings_momentum([3, 2, 2.5, 4, 3.5], 0.8),
            "STRONG",
        )
        self.assertEqual(
            ai_funding_burden(
                capex_growth_pct=[35, 40, 30, 45, 33],
                fcf_after_capex_margin_change_pp=[-6, -7, -5, -8, -6],
                debt_issuance_percentile=85,
                net_debt_rising_share=0.8,
                major_debt_issuance=True,
            ),
            "HIGH",
        )

    def test_determinism_and_isolation(self):
        sample = json.loads(
            (
                Path(__file__).parents[1]
                / "examples"
                / "sample_input.json"
            ).read_text(encoding="utf-8")
        )
        first = evaluate_shadow_v04(sample)
        second = evaluate_shadow_v04(sample)
        self.assertEqual(first, second)
        self.assertNotIn("production", first)
        self.assertTrue(first["shadow"]["pre_reopen_quality_blocked"]["blocked"])

    def test_required_schema_fields(self):
        sample = json.loads(
            (
                Path(__file__).parents[1]
                / "examples"
                / "sample_input.json"
            ).read_text(encoding="utf-8")
        )
        out = evaluate_shadow_v04(sample)
        for key in [
            "pre_reopen_state",
            "pre_reopen_quality_blocked",
            "smh_hard_fail",
            "a_share_dynamic_heat",
            "gld_regime_role",
            "macro_regime",
            "momentum_overlay",
            "ai_fundamental_overlay",
            "proposals",
        ]:
            self.assertIn(key, out["shadow"])
        for key in [
            "evidence_for",
            "evidence_against",
            "dominant_driver",
            "invalidation_condition",
            "confidence",
            "data_dependency",
        ]:
            self.assertIn(key, out["interpreter"])


if __name__ == "__main__":
    unittest.main()
