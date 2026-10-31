def sort_by_amount(total, field_names=None):
    """按求和字段（第 2 个字段）降序排序。

    v1.5.0: total 值为 {字段名: 值} 字典，按 field_names[1] 排序。
    兼容旧格式：total 值为数字时直接按数字排序。

    Args:
        total: {分组键: {字段名: 值, ...}} 或 {分组键: 金额}
        field_names: 字段名列表（可选，默认取第二个字段）

    Returns:
        [(分组键, 数据), ...] 降序排列
    """
    if field_names and len(field_names) >= 2:
        sum_field = field_names[1]
        return sorted(
            total.items(),
            key=lambda item: _to_number(item[1].get(sum_field, 0)),
            reverse=True,
        )
    # 兼容旧格式
    return sorted(
        total.items(),
        key=lambda item: _to_number(item[1]),
        reverse=True,
    )


def _to_number(v):
    """安全转数字，转换失败返回 0。"""
    if isinstance(v, (int, float)):
        return v
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0
