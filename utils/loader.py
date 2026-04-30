import sys
import time

def loading(text="Processing", duration=3):
    for i in range(duration):
        sys.stdout.write(f"\r{text}{'.' * (i % 4)}   ")
        sys.stdout.flush()
        time.sleep(0.5)
    print("\rDone!                ")