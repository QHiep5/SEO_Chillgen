"""Compatibility entry point for evidence-driven R001 re-QA.

Use run_reqa_evidence_driven.py for R001-R052. This wrapper has no scoring
logic, so the legacy hardcoded scorer cannot be invoked accidentally.
"""
from run_reqa_evidence_driven import main

if __name__ == "__main__":
    import sys
    if "--revision" not in sys.argv:
        sys.argv[1:1] = ["--revision", "R001"]
    main()
