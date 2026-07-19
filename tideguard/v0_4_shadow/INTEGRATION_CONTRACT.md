# Integration Contract

1. 运行并冻结 v0.3 Production。
2. 构建 v0.4 Shadow 输入。
3. 调用 `evaluate_shadow_v04(payload)`。
4. 只写入 `shadow` 与 `interpreter`。
5. 断言 Production 序列化内容前后完全一致。
6. 历史不足时保留 DATA_INSUFFICIENT，不推测。
7. v0.4 内容只渲染在现有 Shadow Panel 与 Module Interpreter。
