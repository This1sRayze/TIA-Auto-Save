from tkinter import messagebox
import clr
import schedule
import tkinter as tk
from datetime import datetime
from tkinter import HORIZONTAL, ttk
import os

tia = None

class tia_connect(tk.Tk):
    def __init__(self):
        ''' Initial call, define window and elements'''
        super().__init__()

        self.title("Tia Auto Save")
        self.geometry("250x240")
        self.resizable(True, True)
        self.attributes('-topmost', True)

        self.configure(bg="#2E2E2E")
        self.option_add('*TButton*highlightBackground', '#2E2E2E')
        self.option_add('*TButton*highlightColor', '#2E2E2E')
        self.option_add('*TButton*highlightThickness', 0)

        leftframe_ = tk.Frame(self, bg="#2E2E2E")
        leftframe_.grid(row=0, column=0, rowspan=2, padx=10, pady=10, sticky="nswe")
        
        rightframe_ = tk.Frame(self, bg="#2E2E2E")
        rightframe_.grid(row=0, column=1, rowspan=2, padx=10, pady=10, sticky="nswe")

        bottomframe_ = tk.Frame(self, bg="#2E2E2E")
        bottomframe_.grid(row=2, column=0, columnspan=2, padx=10, pady=0)

        tk.Label(leftframe_, text="TIA Portal Version:", fg="white", bg="#2E2E2E")\
            .grid(row=0, column=0, sticky="w", pady=(0,5))
        self.sv_version_input = tk.StringVar()
        self.entry_version = tk.Entry(leftframe_, textvariable=self.sv_version_input)
        self.entry_version.insert(0, "18")
        self.entry_version.grid(row=1, column=0, sticky="ew", pady=(0,10))
        self.sv_version_input.trace('w', self.update_dll_path)

        tk.Label(leftframe_, text="Select Process:", fg="white", bg="#2E2E2E")\
            .grid(row=2, column=0, sticky="w", pady=(0,5))
        self.sv_cb_Avail_Jobs = tk.StringVar()
        self.cb_Avail_Jobs = ttk.Combobox(leftframe_, textvariable=self.sv_cb_Avail_Jobs)
        self.cb_Avail_Jobs.grid(row=3, column=0, sticky="ew", pady=(0,10))
        
        tk.Label(leftframe_, text="Save Interval (minutes):", fg="white", bg="#2E2E2E")\
            .grid(row=4, column=0, sticky="w", pady=(0,5))
        self.iv_spn_spinval = tk.IntVar(value=5)
        self.spn_interval = ttk.Spinbox(
            leftframe_, from_=1, to=100, increment=1, textvariable=self.iv_spn_spinval)
        self.spn_interval.grid(row=5, column=0, sticky="ew", pady=(0,10))

        self.btn_refresh = tk.Button(
            rightframe_, text="Refresh", bg="#4C4C4C", fg="white",
            command=lambda: (self.update_dll_path(), self.refresh()))
        self.btn_refresh.grid(row=0, column=0, sticky="ew", pady=(0,10))

        self.btn_Man_Save = tk.Button(
            rightframe_, text="Start Saving", bg="green", fg="white",
            command=self.btn_start_toggle)
        self.btn_Man_Save.grid(row=1, column=0, sticky="ew", pady=(0,10))

        self.pb_time_left = ttk.Progressbar(
            bottomframe_, orient=HORIZONTAL, length=200)
        self.pb_time_left.grid(row=0, column=0, sticky="ew", pady=(0,5))
        self.lbl_time_left = tk.Label(
            bottomframe_, text="Next Save in -- min and -- sec.",
            bg="#2E2E2E", fg="white")
        self.lbl_time_left.grid(row=1, column=0, sticky="w")

        self.traces()
        self.first_cycle()
        self.after(0, self.parallel_loop)

    def update_dll_path(self, *args):
        """Load the correct Siemens.Engineering.dll based on version."""
        version = self.sv_version_input.get().strip()
        if not version:
            return
        dll_path = os.path.join(
            f"C:\\Program Files\\Siemens\\Automation\\Portal V{version}",
            "PublicAPI",
            f"V{version}",
            "Siemens.Engineering.dll"
        )
        try:
            clr.AddReference(dll_path)
            global tia
            import Siemens.Engineering as tia
            self.log_message(f"Loaded DLL: {dll_path}")
        except Exception as e:
            messagebox.showerror(
                "DLL Load Failed",
                f"Could not load TIA API DLL from:\n{dll_path}\n\nError: {e}"
            )

    def refresh(self):
        """Only attempt to find processes if TIA module is loaded."""
        if tia is None:
            messagebox.showwarning(
                "TIA Not Loaded",
                "Please enter your Portal version and click Refresh first."
            )
            return
        self.get_processes()
        self.get_job_info()

    def parallel_loop(self):
        loop_interval_ms = 250
        if getattr(self, 'sched_enabled', False):
            schedule.run_pending()
            self.pb_time_left['value'] += loop_interval_ms / 1000
            pct = int((self.pb_time_left['value'] / self.interval_sec) * 100)
            self.title(f"Tia Auto Save - {pct}%")
            rem = (self.interval_sec - self.pb_time_left['value']) / 60
            m, s = divmod(rem * 60, 60)
            self.lbl_time_left.config(
                text=f"Next Save in {int(m)} minutes and {int(s)} seconds.")
            self.btn_Man_Save.config(bg="red", text="Stop Saving")
        else:
            self.pb_time_left['value'] = 0
            self.title("Tia Auto Save")
            self.lbl_time_left.config(text="Next Save in -- min and -- sec.")
            self.btn_Man_Save.config(bg="green", text="Start Saving")

        self.after(loop_interval_ms, self.parallel_loop)

    def first_cycle(self):
        self.sched_enabled = False
        self.iv_spn_spinval.set(5)

    def save_project(self):
        now_str = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        self.pb_time_left['value'] = 0
        try:
            self.myproject.Save()
            self.log_message(f"{now_str} - Saved.")
        except Exception as e:
            msg = f"{now_str} - Save Failed: {e}"
            self.log_message(msg)
            if "EngineeringTargetInvocationException" in str(type(e)):
                messagebox.showwarning(
                    "TIA Auto Save",
                    "Cannot save while ONLINE. Go offline and click OK."
                )

    def log_message(self, message):
        print(message)

    def traces(self):
        self.iv_spn_spinval.trace('w', self.set_save_interval)
        self.sv_cb_Avail_Jobs.trace('w', self.set_job_selection)

    def btn_start_toggle(self):
        self.sched_enabled = not self.sched_enabled

    def get_processes(self):
        self.processes = tia.TiaPortal.GetProcesses()
        self.process_found = len(self.processes) > 0

    def get_job_info(self):
        self.jobs = []
        self.paths = []
        if self.process_found:
            for proc in self.processes:
                p = str(proc.ProjectPath).replace("\\", "\\\\")
                self.paths.append(p)
                self.jobs.append(os.path.basename(p))
            self.cb_Avail_Jobs['values'] = self.jobs
            self.cb_Avail_Jobs.current(0)
        else:
            self.cb_Avail_Jobs['values'] = ['No Processes Found']
            self.cb_Avail_Jobs.current(0)

    def set_job_selection(self, *args):
        if not getattr(self, 'process_found', False):
            return
        sel = self.sv_cb_Avail_Jobs.get()
        idx = self.jobs.index(sel)
        self.sched_enabled = False
        self.mytia = self.processes[idx].Attach()
        self.myproject = self.mytia.Projects[0]

    def set_save_interval(self, *args):
        i = max(1, int(self.iv_spn_spinval.get()))
        self.iv_spn_spinval.set(i)
        self.interval_sec = i * 60
        self.pb_time_left.config(maximum=self.interval_sec)
        self.pb_time_left['value'] = 0
        schedule.clear()
        schedule.every(i).minutes.do(self.save_project)


if __name__ == "__main__":
    app = tia_connect()
    app.mainloop()
