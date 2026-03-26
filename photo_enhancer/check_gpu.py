#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GPU check - called from install.bat"""
import sys

try:
    import torch
    print("  PyTorch version :", torch.__version__)
    if torch.cuda.is_available():
        name = torch.cuda.get_device_name(0)
        mem  = torch.cuda.get_device_properties(0).total_memory / 1024**3
        print("  CUDA            : Available")
        print("  GPU             :", name)
        print("  VRAM            : {:.1f} GB".format(mem))
    else:
        print("  CUDA            : NOT available (CPU mode)")
        print("  Hint: re-run install.bat or check NVIDIA driver")
except ImportError:
    print("  [ERROR] PyTorch is not installed")
    sys.exit(1)
