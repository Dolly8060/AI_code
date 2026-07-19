# TideGuard v0.4 Shadow Test Report

## Result

- Test command: `python -m unittest discover -s tests -v`
- Tests: 9
- Failures: 0
- Errors: 0
- Status: PASS

## Coverage

1. Plumbing Liquidity state thresholds
2. Funding Cost percentile state
3. Duration Supply / Duration Cost composite z-state
4. Risk Transmission ordering
5. PRE_REOPEN Quality Block
6. Momentum slope and acceleration
7. AI Earnings Momentum and Funding Burden
8. Deterministic identical-input output
9. Production isolation and required schema fields

The GitHub Actions workflow reruns the same suite on push and pull request changes under `tideguard/v0_4_shadow/**`.
