#!/usr/bin/env python
"""Visual test for spinner."""
import time
import sys
from src.office_janitor import spinner

print("Starting spinner test...")
print("=" * 40)

spinner.start_spinner_thread()
spinner.set_task("Loading data")

print("Watch the spinner below (5 seconds):")
sys.stdout.flush()
time.sleep(5)

spinner.set_task("Processing files")
print("\nTask changed, watch for 3 more seconds:")
sys.stdout.flush()
time.sleep(3)

# Test with interleaved output
spinner.set_task("Logging test")
for i in range(5):
    spinner.spinner_print(f"  Log message {i+1}")
    time.sleep(0.5)

spinner.clear_task()
spinner.stop_spinner_thread()

print("\n" + "=" * 40)
print("Test complete!")
