from format_checker.report import Violation, render_report


def test_report_lists_original_text_problem_current_and_expected():
    violation = Violation(
        location="第 3 段",
        original_text="摘要",
        problem="字号不符合要求",
        current="三号",
        expected="小二号",
    )

    report = render_report([violation])

    assert "1. 第 3 段" in report
    assert "原文：摘要" in report
    assert "问题：字号不符合要求" in report
    assert "当前格式：三号" in report
    assert "应为格式：小二号" in report


def test_report_shows_success_when_no_violations():
    report = render_report([])

    assert report == "未发现可解析范围内的格式问题。"
