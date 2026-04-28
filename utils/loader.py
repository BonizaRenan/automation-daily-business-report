import sys
import time

def loading(text="Processing", duration=3, done_text=None):
    for i in range(duration):
        sys.stdout.write(f"\r{text}{'.' * (i % 4)}   ")
        sys.stdout.flush()
        time.sleep(0.5)

    if done_text:
        print(f"\r{done_text}                ")