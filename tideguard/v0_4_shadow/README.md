# TideGuard v0.4 Shadow Regime Overlay

独立 Shadow 覆盖层，Production v0.3 不变。

新增：Plumbing Liquidity、Funding Cost、Duration Supply、Duration Cost Gate、Risk Transmission、Fast/Slow Regime、PRE_REOPEN Quality Block、Momentum Overlay、AI Earnings/Funding Burden、双向 Module Interpreter。

运行测试：

```bash
cd tideguard/v0_4_shadow
python -m unittest discover -s tests -v
```

集成时必须先冻结 Production JSON，再调用 `evaluate_shadow_v04()`，只合并 `shadow` 与 `interpreter`。
