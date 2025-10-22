import threading
import itertools
import sys
import time
import multiprocessing
import string
import math
import win32com.client
import os

import pywintypes
# word_file = r'C:\Users\lost\Downloads\114尾牙活動流程表 (新).docx'
word_file = r'C:\Users\lost\Downloads\test.docx'



def try_combinations(charset_slice, found_event):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    chars = string.ascii_uppercase + string.ascii_lowercase + string.digits

    for prefix in charset_slice:  # each process gets a unique starting character
        for combo in itertools.product(chars, repeat=3):  # total length = 10
            if found_event.is_set():
                return

            combination = prefix + ''.join(combo)
            # print(f"Trying {combination}")

            try:
                wb = word.Documents.Open(word_file, False, False, None, combination)
                wb.Unprotect(combination)
                print(f"\n[PID {os.getpid()}] ✅ Success: {combination}", flush=True)
                wb.Close(False)
                word.Quit()
                found_event.set()  # Notify others

                return
            except pywintypes.com_error as e:
                continue

        # optional: remove for speed
        # print(combination)


if __name__ == '__main__':
    cpu_count = multiprocessing.cpu_count()
    found_event = multiprocessing.Event()

    full_charset = string.ascii_uppercase + string.ascii_lowercase + string.digits
    chunk_size = math.ceil(len(full_charset) / cpu_count)
    charset_slices = [full_charset[i:i + chunk_size] for i in range(0, len(full_charset), chunk_size)]





    processes = []
    for charset_slice in charset_slices:
        p = multiprocessing.Process(target=try_combinations, args=(charset_slice, found_event))
        p.start()
        processes.append(p)

    found_event.wait()

    # 🚨 When found_event is set, the main process continues.
    # Now, forcefully terminate the *other* processes that are still running.
    for p in processes:
        if p.is_alive():
            try:
                p.terminate() # 🛑 Forcefully stops the process
                p.join(timeout=1) # Give it a moment to terminate gracefully
            except Exception:
                pass # Ignore errors during termination

    # Now the main process can exit cleanly
    print("Exiting main program.")


