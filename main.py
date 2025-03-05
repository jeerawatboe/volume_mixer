import tkinter as tk
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioSessionManager2

def get_audio_sessions():
    """ดึงรายชื่อโปรแกรมที่ใช้เสียงอยู่"""
    sessions = AudioUtilities.GetAllSessions()
    return {s.Process.name(): s for s in sessions if s.Process}

def set_app_volume(app_name, volume_level):
    """ตั้งค่าเสียงสำหรับโปรแกรมที่กำหนด"""
    sessions = get_audio_sessions()
    if app_name in sessions:
        volume_interface = sessions[app_name].SimpleAudioVolume
        volume_interface.SetMasterVolume(volume_level / 100, None)

# GUI
def update_sliders():
    """อัปเดต Slider ตามโปรแกรมที่กำลังใช้เสียง"""
    for widget in frame.winfo_children():
        widget.destroy()

    audio_sessions = get_audio_sessions()
    for app in audio_sessions:
        frame_row = tk.Frame(frame)
        frame_row.pack(fill="x", padx=10, pady=5)

        label = tk.Label(frame_row, text=app, width=20, anchor="w")
        label.pack(side="left")

        slider = tk.Scale(frame_row, from_=0, to=100, orient="horizontal",
                          command=lambda val, app=app: set_app_volume(app, int(val)))
        slider.set(int(audio_sessions[app].SimpleAudioVolume.GetMasterVolume() * 100))
        slider.pack(side="right", fill="x", expand=True)

    # อัปเดตรายการใหม่ทุก 3 วินาที
    root.after(3000, update_sliders)

# สร้าง GUI
root = tk.Tk()
root.title("Per-App Volume Mixer")

canvas = tk.Canvas(root)
frame = tk.Frame(canvas)
scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")
canvas.create_window((0, 0), window=frame, anchor="nw")

frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# อัปเดตรายชื่อโปรแกรมอัตโนมัติ
update_sliders()

root.mainloop()
