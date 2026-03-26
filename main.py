#!/usr/bin/env python3
"""
DrawingAI Pro - Main Entry Point
=================================

Wrapper למעבר הדרגתי לארכיטקטורה חדשה
כרגע משתמש בקוד הקיים customer_extractor_v3_dual.py
"""

__version__ = "3.1.0"

import sys
from pathlib import Path

# Import existing code
from customer_extractor_v3_dual import main as legacy_main


def main():
    """
    נקודת כניסה ראשית
    
    כרגע מפנה לקוד הקיים
    בעתיד: ישתמש בארכיטקטורה החדשה
    """
    from src.utils.logger import setup_logging
    from pathlib import Path
    setup_logging(log_level="INFO", log_dir=Path("logs"))

    print("╔═══════════════════════════════════════════════════════════════════╗")
    print(f"║   Green Coat — DrawingAI Pro v{__version__} - Smart Blueprint Extractor  ║")
    print("║   📦 Hybrid Mode: new structure + existing engine               ║")
    print("╚═══════════════════════════════════════════════════════════════════╝")
    print()
    
    # הפעלת הקוד הקיים
    legacy_main(folder=sys.argv[1] if len(sys.argv) > 1 else None)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n🛑 Stopped by user")
        sys.exit(0)
    except Exception as e:
        print(f"\n\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
