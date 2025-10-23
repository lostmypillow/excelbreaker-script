import multiprocessing
import os
import sys
import threading
import time

import win32com.client

# --- Configuration ---

# NOTE: The path below is specific to your local machine.

# This script requires a protected Excel file at this location to run properly.

LOCKED_FILE = r'C:\Users\lost\Downloads\114尾牙活動流程表 (新).docx'

TOTAL_PASSWORDS = 1000000  # "0000" to "9999"


# ---------------------


def progress_bar(progress_done_event, current_attempts_shared, total_passwords, found_event):
    """

    Displays a real-time progress bar based on the shared counter.

    Runs in a separate thread.

    """

    bar_length = 50

    while not progress_done_event.is_set():

        # Safely read the current attempt count from the shared value

        with current_attempts_shared.get_lock():

            count = current_attempts_shared.value

        # Ensure count doesn't exceed total_passwords for clean display

        if count > total_passwords:
            count = total_passwords

        # Calculate progress metrics

        percent = (count / total_passwords) * 100

        filled_length = int(bar_length * count // total_passwords)

        # Build the progress bar string

        bar = '█' * filled_length + '-' * (bar_length - filled_length)

        # Output the progress bar using \r to overwrite the line

        sys.stdout.write(

            f'\rBrute-forcing: |{bar}| {percent:6.2f}% ({count}/{total_passwords})'

        )

        sys.stdout.flush()

        # Check if all attempts are done (handles cases where the progress bar thread outlives the workers)

        if count >= total_passwords and not found_event.is_set():
            break  # Exit the loop if all work is complete

        time.sleep(0.1)

    # --- Final Cleanup Message ---

    # Overwrite the bar with a final status message

    if found_event.is_set():

        # Success message is typically printed by try_passwords first

        sys.stdout.write('\r✅ Password found. Stopping processes.' + ' ' * (bar_length + 30) + '\n')

    elif not found_event.is_set() and current_attempts_shared.value >= total_passwords:

        sys.stdout.write('\r❌ All attempts exhausted. Password not found.' + ' ' * (bar_length + 30) + '\n')


def try_passwords(start, end, found_event, current_attempts_shared, success_lock):
    office_app = win32com.client.Dispatch("Word.Application")
    office_app.Visible = False
    for p in range(start, end):

        if found_event.is_set():
            break  # Stop if another process succeeded

        password = f"{p:06d}"  # Convert 7 → "0007", 45 → "0045", etc.

        try:

            # Attempt to open the workbook with the password (blocking, time-consuming operation)

            wb = office_app.Documents.Open(LOCKED_FILE,
                                           False, True,
                                           None, password)

            # --- SUCCESS PATH START ---

            # Acquire the lock. Only one process can enter the critical section.

            with success_lock:

                # After acquiring the lock, check the event AGAIN. The first process

                # to acquire the lock will find this FALSE and become the official winner.

                if not found_event.is_set():
                    # We are the first winner. Print success and set the event.

                    print(f"\n[PID {os.getpid()}] ✅ Success: {password}")

                    # Set the event to immediately stop all other workers' loops.

                    found_event.set()

                # If found_event.is_set() is TRUE here, it means another process

                # beat us to setting the event while we were blocking on the lock.

            # Cleanly close and quit the COM object (must be done by the winning process too)

            wb.Close(False)

            office_app.Quit()

            return  # Exit the worker process



        except Exception as e:

            # On failure, increment the shared counter

            with current_attempts_shared.get_lock():

                current_attempts_shared.value += 1

            # print(f"Error for {password}: {e}") # Uncomment for debugging

            continue

    # Cleanup if the range finishes without finding the password

    office_app.Quit()


if __name__ == '__main__':

    # Initialize shared variables for inter-process communication
    found_event = multiprocessing.Event()
    # NEW: Lock to ensure only one process prints the success message and sets the event.
    success_lock = multiprocessing.Lock()
    # 'i' for signed integer, initialized to 0. This tracks total failures.
    current_attempts = multiprocessing.Value('i', 0)
    cpu_count = multiprocessing.cpu_count()
    ranges = []

    # Calculate ranges for multiprocessing

    chunk_size = TOTAL_PASSWORDS // cpu_count

    start = 0

    for i in range(cpu_count):

        end = start + chunk_size

        if i == cpu_count - 1:
            end = TOTAL_PASSWORDS  # Make sure the last chunk includes all remaining passwords

        ranges.append((start, end))

        start = end

    # 1. Start the progress bar thread

    progress_done_event = threading.Event()

    t = threading.Thread(

        target=progress_bar,

        args=(progress_done_event, current_attempts, TOTAL_PASSWORDS, found_event)

    )

    t.start()

    # 2. Start worker processes

    processes = []

    for start, end in ranges:
        p = multiprocessing.Process(

            target=try_passwords,

            # Pass the new success_lock to the worker function

            args=(start, end, found_event, current_attempts, success_lock)

        )

        p.start()

        processes.append(p)

    # 3. Wait for all processes to finish or for the password to be found

    # Using found_event.wait() is better than joining immediately, as it lets

    # us stop quickly if a process finds the password.

    # We wait up to 1 second in a loop to periodically check the event

    while any(p.is_alive() for p in processes) and not found_event.is_set():
        time.sleep(1)

    # If the password was found, terminate all other processes immediately

    if found_event.is_set():

        for p in processes:

            if p.is_alive():
                p.terminate()

    # Ensure all processes are joined

    for p in processes:
        p.join()

    # 4. Stop the progress bar thread and wait for it to clean up

    progress_done_event.set()

    t.join()
