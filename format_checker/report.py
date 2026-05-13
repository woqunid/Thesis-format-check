from dataclasses import dataclass


@dataclass(frozen=True)
class Violation:
    location: str
    original_text: str
    problem: str
    current: str
    expected: str


def render_report(violations: list[Violation]) -> str:
    if not violations:
        return "未发现可解析范围内的格式问题。"

    blocks = [_render_violation(index, item) for index, item in enumerate(violations, 1)]
    return "\n\n".join(blocks)


def _render_violation(index: int, violation: Violation) -> str:
    return "\n".join(
        [
            f"{index}. {violation.location}",
            f"原文：{violation.original_text}",
            f"问题：{violation.problem}",
            f"当前格式：{violation.current}",
            f"应为格式：{violation.expected}",
        ]
    )
