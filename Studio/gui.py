import customtkinter as ctk
import subprocess
import threading
import os
import re
from datetime import datetime
import queue

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class StudioAnalyzer(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Studio Finance Analyzer")
        self.geometry("1200x800")
        
        self.studio_var = ctk.StringVar(value="Hohlovka")
        studio_frame = ctk.CTkFrame(self)
        studio_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(studio_frame, text="Выберите студию:").pack(side="left", padx=10)
        ctk.CTkOptionMenu(studio_frame, values=["Hohlovka", "Yauza"], variable=self.studio_var).pack(side="left", padx=10)
        
        input_frame = ctk.CTkFrame(self)
        input_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(input_frame, text="Период (dd mm yyyy / mm yyyy / dd-dd mm yyyy):").pack(side="left", padx=10)
        self.period_entry = ctk.CTkEntry(input_frame, width=200)
        self.period_entry.pack(side="left", padx=10)
        
        self.start_btn = ctk.CTkButton(input_frame, text="Запустить анализ", command=self.start_analysis)
        self.start_btn.pack(side="left", padx=10)
        
        self.progress = ctk.CTkProgressBar(self)
        self.progress.pack(pady=10, padx=10, fill="x")
        self.progress.set(0)
        
        self.log_text = ctk.CTkTextbox(self, height=200)
        self.log_text.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, padx=10, fill="both", expand=True)
        self.tree = ctk.CTkScrollableFrame(self.tree_frame)
        self.tree.pack(fill="both", expand=True)
        
        self.confirm_frame = ctk.CTkFrame(self)
        self.yes_btn = ctk.CTkButton(self.confirm_frame, text="Да, внести в таблицу", command=lambda: self.send_input("yes"))
        self.no_btn = ctk.CTkButton(self.confirm_frame, text="Нет, отменить", command=lambda: self.send_input("no"))
        self.yes_btn.pack(side="left", padx=10)
        self.no_btn.pack(side="left", padx=10)
        self.confirm_frame.pack(pady=10, padx=10, fill="x")
        self.confirm_frame.pack_forget()
        
        self.process = None
        self.output_queue = queue.Queue()
        self.waiting_confirm = False
        
        self.animate_progress()
    
    def log(self, text, color="white"):
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        if "Day " in text:
            self.log_text.tag_add("day", "end-2l", "end")
            self.log_text.tag_config("day", foreground="cyan")
        elif "руб." in text:
            self.log_text.tag_add("money", "end-2l", "end")
            self.log_text.tag_config("money", foreground="gold")
        elif any(cat in text for cat in ["Фотосъемка", "Банкет", "Видео", "Мероприятие"]):
            self.log_text.tag_add("category", "end-2l", "end")
            self.log_text.tag_config("category", foreground="lightgreen")
    
    def animate_progress(self):
        if self.progress.get() < 1:
            val = self.progress.get() + 0.01
            self.progress.set(min(val, 0.9))
            self.after(200, self.animate_progress)
    
    def start_analysis(self):
        period = self.period_entry.get().strip()
        if not period:
            self.log("Ошибка: Введите период!", "red")
            return
        
        studio = self.studio_var.get()
        folder = studio
        
        self.log_text.delete("1.0", "end")
        for widget in self.tree.winfo_children():
            widget.destroy()
        self.progress.set(0)
        self.confirm_frame.pack_forget()
        self.waiting_confirm = False
        
        self.log(f"Запуск для {studio}, период: {period}")
        
        def run_script():
            try:
                os.chdir(os.path.join(os.path.dirname(__file__), folder))
                main_file = "main.py" if studio == "Hohlovka" else "yauza_main.py"
                cmd = ["python", main_file]
                self.process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    bufsize=1,
                    universal_newlines=True
                )
                
                self.process.stdin.write(period + "\n")
                self.process.stdin.flush()
                
                buffer = ""
                while True:
                    char = self.process.stdout.read(1)
                    if char == "" and self.process.poll() is not None:
                        if buffer:
                            self.output_queue.put(buffer.strip())
                        break
                    if char:
                        buffer += char
                        if "\n" in char:
                            self.output_queue.put(buffer.strip())
                            buffer = ""
                        elif "Внести в таблицу? (yes/no): " in buffer:  # Детект промпта без \n
                            self.output_queue.put(buffer.strip())
                            buffer = ""
                
                self.output_queue.put("END")
            except Exception as e:
                self.output_queue.put(f"Ошибка запуска: {e}")
        
        threading.Thread(target=run_script, daemon=True).start()
        
        def process_output():
            try:
                line = self.output_queue.get_nowait()
                if line == "END":
                    self.progress.set(1)
                    self.log("Анализ завершён!")
                    self.parse_table()
                    return
                self.log(line)
                
                if "Внести в таблицу? (yes/no):" in line:
                    self.waiting_confirm = True
                    self.show_confirm()
            except queue.Empty:
                pass
            self.after(100, process_output)
        
        process_output()
    
    def show_confirm(self):
        if self.waiting_confirm:
            self.confirm_frame.pack(pady=10, padx=10, fill="x")
    
    def send_input(self, choice):
        if self.process and self.waiting_confirm:
            self.process.stdin.write(choice + "\n")
            self.process.stdin.flush()
            self.log(f"Выбор: {choice}")
            self.confirm_frame.pack_forget()
            self.waiting_confirm = False
    
    def parse_table(self):
        log_content = self.log_text.get("1.0", "end")
        days = re.findall(r'Day (\d{2}\.\d{2}\.\d{4}):(.*?)(?=Day |Общие часы|$)', log_content, re.DOTALL | re.UNICODE)
        
        row = 0
        for day, content in days:
            ctk.CTkLabel(self.tree, text=f"День: {day}", font=ctk.CTkFont(size=14, weight="bold")).grid(row=row, column=0, columnspan=3, pady=5, sticky="w")
            row += 1
            
            cats = re.findall(r'([А-Яа-я /]+): (\d+\.?\d*) ч \(бронирований: (\d+)\)', content, re.UNICODE)
            for cat, hours, count in cats:
                ctk.CTkLabel(self.tree, text=cat).grid(row=row, column=0, sticky="w")
                ctk.CTkLabel(self.tree, text=f"{hours} ч").grid(row=row, column=1, sticky="w")
                ctk.CTkLabel(self.tree, text=f"Броней: {count}").grid(row=row, column=2, sticky="w")
                row += 1
            
            money_matches = re.findall(r'(Предоплаты фото|Предоплаты видео|По факту фото|По факту видео|Доп. услуги|Парковки сумма|Школа по часам): (\d+) руб\.', content)
            for label, amount in money_matches:
                ctk.CTkLabel(self.tree, text=label).grid(row=row, column=0, sticky="w")
                ctk.CTkLabel(self.tree, text=f"{amount} руб.").grid(row=row, column=1, columnspan=2, sticky="w")
                row += 1
        
        total_match = re.search(r'Общие часы за период:(.*)', log_content, re.DOTALL | re.UNICODE)
        if total_match:
            ctk.CTkLabel(self.tree, text="Общие итоги:", font=ctk.CTkFont(size=14, weight="bold")).grid(row=row, column=0, columnspan=3, pady=10, sticky="w")
            row += 1
            totals = re.findall(r'([А-Яа-я /]+): (\d+\.?\d*) ч', total_match.group(1), re.UNICODE)
            for cat, h in totals:
                ctk.CTkLabel(self.tree, text=f"{cat}: {h} ч").grid(row=row, column=0, columnspan=3, sticky="w")
                row += 1

if __name__ == "__main__":
    app = StudioAnalyzer()
    app.mainloop()