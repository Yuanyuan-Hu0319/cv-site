#!/usr/bin/env python3
"""Build the optimized extended CV from cv-data.json."""

from build_optimized_cvs import build_extended, load_data


if __name__ == "__main__":
    build_extended(load_data())
