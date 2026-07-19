from __future__ import annotations

from statistics import mean, median
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence


def _valid(values: Iterable[Optional[float]]) -> List[float]:
    return [float(v) for v in values if v is not None]


def plumbing_liquidity_state(
    sofr_iorb_bp: Optional[float],
    srf_usage_usd: Optional[float],
    reserve_4w_change_pct: Optional[float],
) -> Dict[str, Any]:
    states: List[str] = []
    if sofr_iorb_bp is not None:
        states.append("STRESS" if sofr_iorb_bp > 10 else "WATCH" if sofr_iorb_bp > 5 else "NORMAL")
    if srf_usage_usd is not None:
        states.append(
            "STRESS"
            if srf_usage_usd > 10_000_000_000
            else "WATCH"
            if srf_usage_usd > 1_000_000_000
            else "NORMAL"
        )
    if reserve_4w_change_pct is not None:
        states.append(
            "STRESS"
            if reserve_4w_change_pct < -5
            else "WATCH"
            if reserve_4w_change_pct < -3
            else "NORMAL"
        )

    if len(states) < 2:
        state = "DATA_INSUFFICIENT"
    elif "STRESS" in states or states.count("WATCH") >= 2:
        state = "STRESS"
    elif states.count("WATCH") == 1:
        state = "WATCH"
    else:
        state = "NORMAL"

    return {
        "state": state,
        "sofr_iorb_bp": sofr_iorb_bp,
        "srf_usage_usd": srf_usage_usd,
        "reserve_4w_change_pct": reserve_4w_change_pct,
        "valid_inputs": len(states),
    }


def funding_cost_state(percentiles: Mapping[str, Optional[float]]) -> Dict[str, Any]:
    values = _valid(percentiles.values())
    if len(values) < 3:
        state = "DATA_INSUFFICIENT"
    else:
        real10 = percentiles.get("tips10y")
        if sum(v >= 90 for v in values) >= 3 or (real10 is not None and real10 >= 95):
            state = "EXTREME"
        elif sum(v >= 80 for v in values) >= 2 or (real10 is not None and real10 >= 90):
            state = "HIGH"
        elif sum(v <= 20 for v in values) >= 3:
            state = "LOW"
        else:
            state = "NORMAL"
    return {"state": state, "percentiles": dict(percentiles), "valid_inputs": len(values)}


def composite_z_state(z_values: Mapping[str, Optional[float]]) -> Dict[str, Any]:
    values = _valid(z_values.values())
    if len(values) < 3:
        return {
            "state": "DATA_INSUFFICIENT",
            "composite_z": None,
            "valid_inputs": len(values),
            "components": dict(z_values),
        }

    score = mean(values)
    state = (
        "LOW"
        if score <= -0.75
        else "NORMAL"
        if score < 0.75
        else "HIGH"
        if score < 1.5
        else "EXTREME"
    )
    return {
        "state": state,
        "composite_z": round(score, 4),
        "valid_inputs": len(values),
        "components": dict(z_values),
    }


def risk_transmission_state(
    *,
    red_tail: bool,
    vix: Optional[float],
    hy_oas_5d_change_bp: Optional[float],
    hy_oas_2d_change_bp: Optional[float],
    ai_semi_state: Optional[str],
    repair_pass_count: Optional[int],
    a_share_state: Optional[str],
    funding_cost: Optional[str],
    duration_cost: Optional[str],
) -> Dict[str, Any]:
    if sum(
        x is not None
        for x in [vix, hy_oas_5d_change_bp, ai_semi_state, repair_pass_count, a_share_state]
    ) < 3:
        return {"state": "DATA_INSUFFICIENT", "evidence": ["关键输入不足"]}

    if red_tail or (
        vix is not None
        and vix >= 28
        and hy_oas_5d_change_bp is not None
        and hy_oas_5d_change_bp >= 20
    ):
        return {"state": "BROAD_SYSTEMIC", "evidence": ["系统性阈值成立"]}

    if (
        ai_semi_state in {"FAIL", "HARD_FAIL"}
        or (repair_pass_count is not None and repair_pass_count <= 2)
        or a_share_state in {"WARNING", "A_SHARE_STRUCTURE_RISK"}
    ):
        return {"state": "RATES_TO_EQUITY", "evidence": ["压力已传至股票层"]}

    if (
        hy_oas_5d_change_bp is not None
        and hy_oas_5d_change_bp >= 10
        or hy_oas_2d_change_bp is not None
        and hy_oas_2d_change_bp > 5
    ):
        return {"state": "RATES_TO_CREDIT", "evidence": ["信用利差达到传导阈值"]}

    if funding_cost in {"HIGH", "EXTREME"} or duration_cost in {"HIGH", "EXTREME"}:
        return {"state": "RATES_ONLY", "evidence": ["高资金成本尚未跨层传导"]}

    return {"state": "CONTAINED", "evidence": ["压力未跨层传导"]}


def slow_regime(duration_cost_state_value: str) -> str:
    return {
        "LOW": "LOW_DURATION_COST",
        "NORMAL": "NORMAL_DURATION_COST",
        "HIGH": "HIGH_DURATION_COST",
        "EXTREME": "EXTREME_DURATION_COST",
    }.get(duration_cost_state_value, "DATA_INSUFFICIENT")


def pre_reopen_quality_blocked(
    pre_reopen_state: str,
    duration_cost_state_value: str,
    smh_hard_fail: bool,
    plumbing_state: str,
    a_share_dynamic_state: Optional[str],
    a_share_structure_risk: bool,
) -> Dict[str, Any]:
    if pre_reopen_state not in {"PRE_REOPEN_LOW", "PRE_REOPEN_HIGH"}:
        return {"blocked": False, "reasons": []}

    reasons: List[str] = []
    if duration_cost_state_value in {"HIGH", "EXTREME"}:
        reasons.append("DURATION_COST_HIGH")
    if smh_hard_fail:
        reasons.append("SMH_HARD_FAIL")
    if plumbing_state == "STRESS":
        reasons.append("PLUMBING_STRESS")
    if a_share_dynamic_state in {"HIGH", "EXTREME", "WARNING"}:
        reasons.append("A_SHARE_DYNAMIC_HEAT")
    if a_share_structure_risk:
        reasons.append("A_SHARE_STRUCTURE_RISK")
    return {"blocked": bool(reasons), "reasons": reasons}


def slope_label(change: Optional[float], noise_band: float) -> str:
    if change is None:
        return "DATA_INSUFFICIENT"
    if change >= noise_band:
        return "RISING"
    if change <= -noise_band:
        return "FALLING"
    return "FLAT"


def acceleration_label(
    latest_1d_change: Optional[float],
    previous_2d_average_change: Optional[float],
    noise_band: float,
) -> str:
    if latest_1d_change is None or previous_2d_average_change is None:
        return "DATA_INSUFFICIENT"
    delta = latest_1d_change - previous_2d_average_change
    if delta >= noise_band:
        return "ACCELERATING"
    if delta <= -noise_band:
        return "DECELERATING"
    return "FLAT"


def ai_earnings_momentum(
    revisions_pct: Sequence[float], positive_share: Optional[float]
) -> str:
    if len(revisions_pct) < 5 or positive_share is None:
        return "DATA_INSUFFICIENT"
    med = median(revisions_pct)
    if med >= 2 and positive_share >= 0.60:
        return "STRONG"
    if med <= -2 or positive_share <= 0.40:
        return "WEAK"
    return "NEUTRAL"


def ai_funding_burden(
    *,
    capex_growth_pct: Sequence[float],
    fcf_after_capex_margin_change_pp: Sequence[float],
    debt_issuance_percentile: Optional[float],
    net_debt_rising_share: Optional[float],
    major_debt_issuance: Optional[bool],
) -> str:
    if (
        len(capex_growth_pct) < 5
        or len(fcf_after_capex_margin_change_pp) < 5
        or net_debt_rising_share is None
        or major_debt_issuance is None
    ):
        return "DATA_INSUFFICIENT"

    capex_med = median(capex_growth_pct)
    fcf_med = median(fcf_after_capex_margin_change_pp)
    if (capex_med >= 30 and fcf_med <= -5) or (
        debt_issuance_percentile is not None and debt_issuance_percentile >= 80
    ):
        return "HIGH"
    if capex_med >= 15 or net_debt_rising_share >= 0.50:
        return "RISING"
    if capex_med < 10 and fcf_med >= 0 and not major_debt_issuance:
        return "LOW"
    return "NORMAL"


def evaluate_shadow_v04(payload: Mapping[str, Any]) -> Dict[str, Any]:
    plumbing = plumbing_liquidity_state(
        payload.get("sofr_iorb_bp"),
        payload.get("srf_usage_usd"),
        payload.get("reserve_4w_change_pct"),
    )
    funding = funding_cost_state(payload.get("funding_cost_percentiles", {}))
    supply = composite_z_state(payload.get("duration_supply_pressure_z", {}))

    duration_inputs = dict(payload.get("duration_cost_pressure_z", {}))
    if supply["composite_z"] is not None:
        duration_inputs.setdefault("duration_supply", supply["composite_z"])
    duration = composite_z_state(duration_inputs)

    transmission = risk_transmission_state(
        red_tail=bool(payload.get("red_tail", False)),
        vix=payload.get("vix"),
        hy_oas_5d_change_bp=payload.get("hy_oas_5d_change_bp"),
        hy_oas_2d_change_bp=payload.get("hy_oas_2d_change_bp"),
        ai_semi_state=payload.get("ai_semi_state"),
        repair_pass_count=payload.get("repair_pass_count"),
        a_share_state=payload.get("a_share_state"),
        funding_cost=funding["state"],
        duration_cost=duration["state"],
    )

    quality = pre_reopen_quality_blocked(
        payload.get("pre_reopen_state", "NONE"),
        duration["state"],
        bool(payload.get("smh_hard_fail", False)),
        plumbing["state"],
        payload.get("a_share_dynamic_state"),
        bool(payload.get("a_share_structure_risk", False)),
    )

    momentum: Dict[str, Any] = {}
    for name, spec in payload.get("momentum", {}).items():
        momentum[name] = {
            "slope": slope_label(spec.get("change"), float(spec["noise_band"])),
            "acceleration": acceleration_label(
                spec.get("latest_1d_change"),
                spec.get("previous_2d_average_change"),
                float(spec["noise_band"]),
            ),
            "raw": dict(spec),
        }

    ai = payload.get("ai_fundamentals", {})
    return {
        "shadow": {
            "pre_reopen_state": payload.get("pre_reopen_state", "NONE"),
            "pre_reopen_quality_blocked": quality,
            "smh_hard_fail": bool(payload.get("smh_hard_fail", False)),
            "a_share_dynamic_heat": {
                "state": payload.get("a_share_dynamic_state", "DATA_INSUFFICIENT")
            },
            "gld_regime_role": "REGIME_ONLY",
            "macro_regime": {
                "plumbing_liquidity": plumbing,
                "funding_cost": funding,
                "duration_supply": supply,
                "duration_cost_gate": duration,
                "risk_transmission": transmission,
                "slow_regime": {"state": slow_regime(duration["state"])},
                "fast_state": {
                    "total_state": payload.get("production_total_state"),
                    "trend": payload.get("production_trend"),
                },
            },
            "momentum_overlay": momentum,
            "ai_fundamental_overlay": {
                "earnings_momentum": ai_earnings_momentum(
                    ai.get("revisions_pct", []), ai.get("positive_share")
                ),
                "funding_burden": ai_funding_burden(
                    capex_growth_pct=ai.get("capex_growth_pct", []),
                    fcf_after_capex_margin_change_pp=ai.get(
                        "fcf_after_capex_margin_change_pp", []
                    ),
                    debt_issuance_percentile=ai.get("debt_issuance_percentile"),
                    net_debt_rising_share=ai.get("net_debt_rising_share"),
                    major_debt_issuance=ai.get("major_debt_issuance"),
                ),
                "price_confirmation": payload.get("ai_semi_state"),
                "last_update": ai.get("last_update"),
            },
            "proposals": [],
        },
        "interpreter": {
            "evidence_for": list(payload.get("evidence_for", []))[:3],
            "evidence_against": list(payload.get("evidence_against", []))[:3],
            "dominant_driver": payload.get("dominant_driver", "MIXED"),
            "invalidation_condition": list(payload.get("invalidation_condition", [])),
            "confidence": payload.get("confidence", "LOW"),
            "data_dependency": list(payload.get("data_dependency", [])),
        },
    }
