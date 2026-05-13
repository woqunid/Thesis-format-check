from pathlib import Path

from format_checker.rules import check_file


def test_checker_parses_template_docx_without_crashing():
    template = Path("00 学生用-附件1：2026届毕业论文（设计）撰写规范模板(1).docx")

    violations = check_file(template)

    assert isinstance(violations, list)
