# %%
import os
import csv
import argparse
from typing import List, Tuple, Optional, Dict
from dataclasses import dataclass

from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
)
from pptx import Presentation
from pptx.table import Table
from pptx.enum.shapes import MSO_SHAPE_TYPE
from tqdm import tqdm
from dotenv import load_dotenv

load_dotenv()

# OpenAI SDK v1.x
try:
    from openai import OpenAI
except Exception as e:
    OpenAI = None  # handled in main

SLIDE_FILE = "input.pptx"

prs = Presentation(SLIDE_FILE)

# %%
