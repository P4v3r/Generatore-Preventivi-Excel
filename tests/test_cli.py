"""Tests per cli.py."""

import argparse

from cli import parse_args, validate_args


def test_parse_args_defaults():
    args = parse_args(["--input", "test.csv"])
    
    assert args.input == "test.csv"
    assert args.template == "milano"  # Default
    assert args.use_ai == False


def test_parse_args_full():
    args = parse_args([
        "--input", "lavori.csv",
        "--template", "liguria",
        "--use-ai"
    ])
    
    assert args.template == "liguria"
    assert args.use_ai == True


def test_parse_args_verbose():
    args = parse_args(["--input", "lavori.csv", "--verbose"])
    
    assert args.verbose == True


def test_validate_args_missing_file():
    class Args:
        input = "/nonexistent/file.csv"
    
    result = validate_args(Args())
    assert result == False


def test_validate_args_existing_file(tmp_path):
    test_file = tmp_path / "test.csv"
    test_file.write_text("test")
    
    class Args:
        input = str(test_file)
    
    result = validate_args(Args())
    assert result == True