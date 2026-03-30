import ctypes
import os
import subprocess
import time
import wave
import winsound


WAV_PATH = os.path.join(os.path.dirname(__file__), "bad_apple.wav")
ERROR_VBS_PATH = os.path.join(os.path.dirname(__file__), "error.vbs")
WARNING_VBS_PATH = os.path.join(os.path.dirname(__file__), "warning.vbs")
TEMPO = 138
MAX_ACTIVE_POPUPS = 40
VK_ESCAPE = 0x1B


def get_duration_seconds(path: str) -> float:
    with wave.open(path, "rb") as wav_file:
        frames = wav_file.getnframes()
        frame_rate = wav_file.getframerate()
    return frames / float(frame_rate)


def cleanup_processes(processes: list[subprocess.Popen]) -> list[subprocess.Popen]:
    return [process for process in processes if process.poll() is None]


def stop_processes(processes: list[subprocess.Popen]) -> None:
    for process in processes:
        if process.poll() is None:
            process.terminate()

    for process in processes:
        try:
            process.wait(timeout=0.5)
        except subprocess.TimeoutExpired:
            process.kill()


def trigger_popup(script_path: str) -> subprocess.Popen | None:
    if not os.path.exists(script_path):
        return None

    return subprocess.Popen(
        ["wscript.exe", script_path],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def escape_pressed() -> bool:
    return bool(ctypes.windll.user32.GetAsyncKeyState(VK_ESCAPE) & 0x8000)


def main() -> None:
    if not os.path.exists(WAV_PATH):
        print(f"File Not Found: {WAV_PATH}")
        return

    missing_scripts = [
        path for path in (ERROR_VBS_PATH, WARNING_VBS_PATH) if not os.path.exists(path)
    ]
    if missing_scripts:
        print("Missing VBS file:")
        for path in missing_scripts:
            print(path)
        return

    duration = get_duration_seconds(WAV_PATH)
    beat_interval = 60.0 / TEMPO
    scripts = [ERROR_VBS_PATH, WARNING_VBS_PATH]

    print("Playing bad_apple.wav")
    print("Press ESC to stop and exit.")

    winsound.PlaySound(WAV_PATH, winsound.SND_FILENAME | winsound.SND_ASYNC)
    start_time = time.time()
    next_beat_time = start_time
    beat_index = 0
    popup_processes: list[subprocess.Popen] = []

    try:
        while time.time() - start_time < duration:
            now = time.time()
            popup_processes = cleanup_processes(popup_processes)

            while now >= next_beat_time:
                if len(popup_processes) >= MAX_ACTIVE_POPUPS:
                    stop_processes(popup_processes)
                    popup_processes = []

                process = trigger_popup(scripts[beat_index % len(scripts)])
                if process is not None:
                    popup_processes.append(process)

                beat_index += 1
                next_beat_time += beat_interval

            if escape_pressed():
                stop_processes(popup_processes)
                winsound.PlaySound(None, winsound.SND_PURGE)
                print("Stopped.")
                return

            time.sleep(0.05)
        stop_processes(popup_processes)
        print("Finished.")
    finally:
        stop_processes(popup_processes)
        winsound.PlaySound(None, winsound.SND_PURGE)


if __name__ == "__main__":
    main()
