# -*- coding: utf-8 -*-
"""
Release build (locked VBA project).
Use --unviewable only if you really need irreversible protection.
"""
import argparse
from rebuild_xlam import rebuild


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Rebuild XLAM for release (password locked)",
    )
    parser.add_argument(
        "--unviewable",
        action="store_true",
        help="Apply irreversible UNVIEWABLE patch after build",
    )
    args = parser.parse_args()

    rebuild(dev_mode=False, make_unviewable=args.unviewable)
