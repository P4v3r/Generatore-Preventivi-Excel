"""CLI interface per generatore preventivi."""

import argparse
import sys
from pathlib import Path


def parse_args(args=None):
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Genera preventivi Excel da template regionali",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Esempi:
  %(prog)s --input nuovi_lavori.csv --template milano
  %(prog)s --input lavori.csv --template liguria --use-ai
        """,
    )
    
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path al file CSV con nuovi lavori",
    )
    
    parser.add_argument(
        "--template", "-t",
        choices=["milano", "liguria"],
        default="milano",
        help="Template regionale (default: milano)",
    )
    
    parser.add_argument(
        "--use-ai",
        action="store_true",
        help="Attiva fallback AI per match non trovati",
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Output verboso",
    )
    
    return parser.parse_args(args)


def validate_args(args) -> bool:
    """Valida gli argomenti."""
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f"❌ Errore: File non trovato: {input_path}")
        return False
    
    return True


def print_banner():
    """Stampa banner iniziale."""
    print("=" * 50)
    print("  Generatore Preventivi Excel")
    print("=" * 50)
    print()


def print_summary(generati: int, errori: int, log_path: Path):
    """Stampa riepilogo fine esecuzione."""
    print()
    print("=" * 50)
    print(f"  ✅ Preventivi generati: {generati}")
    print(f"  ❌ Errori: {errori}")
    print(f"  📋 Log: {log_path}")
    print("=" * 50)