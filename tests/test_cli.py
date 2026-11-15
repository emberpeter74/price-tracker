from pathlib import Path

import pricetracker.cli as cli
from pricetracker.excel_io import InputError


def test_no_command_prints_usage_with_both_commands_and_returns_2(capsys):
    assert cli.main([]) == 2
    text = capsys.readouterr().out
    assert "scan" in text and "report" in text and "Traceback" not in text


def test_scan_calls_run_scan_with_default_folders(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(cli, "base_dir", lambda: tmp_path)
    monkeypatch.setattr(cli, "run_scan", lambda i, o: calls.append((i, o)) or o / "Prices x.xlsx")
    assert cli.main(["scan"]) == 0
    assert calls == [(tmp_path / "input", tmp_path / "output")]


def test_folders_can_be_overridden(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(cli, "run_scan", lambda i, o: calls.append((i, o)))
    cli.main(["scan", "--input", str(tmp_path / "in"), "--output", str(tmp_path / "out")])
    assert calls == [(tmp_path / "in", tmp_path / "out")]


def test_report_calls_build_report(monkeypatch, tmp_path, capsys):
    monkeypatch.setattr(cli, "base_dir", lambda: tmp_path)
    monkeypatch.setattr(cli, "build_report", lambda o: o / "Price comparison.xlsx")
    assert cli.main(["report"]) == 0
    assert "Price comparison.xlsx" in capsys.readouterr().out


def test_output_folder_is_created_if_missing(monkeypatch, tmp_path):
    monkeypatch.setattr(cli, "base_dir", lambda: tmp_path)
    monkeypatch.setattr(cli, "build_report", lambda o: o / "x.xlsx")
    cli.main(["report"])
    assert (tmp_path / "output").is_dir()


def test_input_error_prints_message_without_traceback_and_returns_1(monkeypatch, tmp_path, capsys):
    def fail(i, o):
        raise InputError("The input file Items.xlsx was not found in C:\\x.")
    monkeypatch.setattr(cli, "base_dir", lambda: tmp_path)
    monkeypatch.setattr(cli, "run_scan", fail)
    assert cli.main(["scan"]) == 1
    text = capsys.readouterr().out
    assert "Items.xlsx was not found" in text and "Traceback" not in text


def test_ctrl_c_returns_130_without_traceback(monkeypatch, tmp_path, capsys):
    def stop(i, o):
        raise KeyboardInterrupt()
    monkeypatch.setattr(cli, "base_dir", lambda: tmp_path)
    monkeypatch.setattr(cli, "run_scan", stop)
    assert cli.main(["scan"]) == 130 and "Traceback" not in capsys.readouterr().out


def test_unexpected_error_prints_one_line_and_writes_error_log(monkeypatch, tmp_path, capsys):
    def boom(i, o):
        raise RuntimeError("kaboom")
    monkeypatch.setattr(cli, "base_dir", lambda: tmp_path)
    monkeypatch.setattr(cli, "run_scan", boom)
    assert cli.main(["scan"]) == 1
    text = capsys.readouterr().out
    assert "error.log" in text and "Traceback" not in text
    log = (tmp_path / "output" / "error.log").read_text(encoding="utf-8")
    assert "kaboom" in log and "Traceback" in log


def test_base_dir_is_the_working_directory_in_development():
    assert cli.base_dir() == Path.cwd()
