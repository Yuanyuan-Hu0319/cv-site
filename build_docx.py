#!/usr/bin/env python3
"""Build the optimized one-page CV from cv-data.json."""

from build_optimized_cvs import build_one_page, load_data


if __name__ == "__main__":
    build_one_page(load_data())
