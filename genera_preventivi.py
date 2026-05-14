#!/usr/bin/env python3
"""Entry point CLI per generatore preventivi."""

from pathlib import Path

from cli import parse_args, validate_args, print_banner, print_summary
from main import PreventiviGenerator


def main():
    args = parse_args()
    
    if not validate_args(args):
        return 1
    
    print_banner()
    
    if args.verbose:
        print(f"📂 Input: {args.input}")
        print(f"📋 Template: {args.template}")
        print(f"🤖 AI: {'Attivo' if args.use_ai else 'Disattivo'}")
        print()
    
    # Crea generatore e run
    generator = PreventiviGenerator(
        regione=args.template,
        use_ai=args.use_ai,
    )
    
    output_files, errors = generator.run(args.input)
    
    # Riepilogo
    print_summary(len(output_files), len(errors), Path("log_matching.csv"))
    
    # Stampa errori se presenti
    if errors:
        print()
        print("⚠️ Errori:")
        for err in errors:
            print(f"  - {err}")
    
    return 0 if not errors else 1


if __name__ == "__main__":
    exit(main())