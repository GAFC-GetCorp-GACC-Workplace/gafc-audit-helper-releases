# -*- coding: utf-8 -*-
"""
Development build (unlocked VBA project).
Uses the same template as release but keeps code viewable.
"""
from rebuild_xlam import rebuild


if __name__ == "__main__":
    rebuild(dev_mode=True, make_unviewable=False)
