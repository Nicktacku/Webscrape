import asyncio
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from selenium_driverless import webdriver
from selenium_driverless.types.by import By

# -----------------------------------------------------------------------------
# Data Loader (Reads and standardizes both Time In and Time Out columns)
# -----------------------------------------------------------------------------
def load_and_normalize_df(file_path, log_callback):
    """Loads the Excel file and detects both Time In and Time Out data columns."""
    df = pd.read_excel(file_path)
    df.columns = [str(col).strip() for col in df.columns]
    
    time_in_col = None
    for possible_time in ['Time In', 'Actual Time In', 'time in', 'TimeIn']:
        if possible_time in df.columns:
            time_in_col = possible_time
            break

    time_out_col = None
    for possible_time_out in ['Time Out', 'Actual Time Out', 'time out', 'TimeOut']:
        if possible_time_out in df.columns:
            time_out_col = possible_time_out
            break

    if not time_in_col and len(df.columns) > 0:
        time_in_col = df.columns[0]
    if not time_out_col and len(df.columns) > 1:
        time_out_col = df.columns[1]

    if not time_in_col:
        log_callback("SETUP ERROR: The Excel file layout is empty or undetected.")
        raise ValueError("Empty File Layout")
    if not time_out_col:
        log_callback("SETUP ERROR: Could not find a second column for Time Out in your file.")
        raise ValueError("Missing Time Out column")

    log_callback(f"-> Reading times from columns: In=[{time_in_col}] | Out=[{time_out_col}]")
    df = df.rename(columns={time_in_col: 'TARGET_TIME_IN', time_out_col: 'TARGET_TIME_OUT'})
    return df

# -----------------------------------------------------------------------------
# Optimized Core Engine (Dynamic Waits & Smart Postbacks)
# -----------------------------------------------------------------------------
async def run_automation(gui_instance, username, password, start_date, end_date, file_path, log_callback):
    try:
        log_callback(f"Opening file: {file_path}")
        df = load_and_normalize_df(file_path, log_callback)
        log_callback(f"Total time entries parsed from Excel: {len(df)}")
    except Exception as e:
        gui_instance.root.after(0, lambda: messagebox.showerror("File Error", f"Failed to parse data sheet layout:\n{str(e)}"))
        return

    log_callback("Launching Chrome browser automated engine in HEADLESS mode...")
    
    # Correct headless configuration for selenium_driverless
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    
    async with webdriver.Chrome(options=options) as driver:
        # 1. Login Step
        await driver.get("https://my-ices.teligent.com.ph/Portal/SignIn.aspx", wait_load=True)
        log_callback("Logging into the portal...")
        await (await driver.find_element(By.ID, "txtCompanyCode", timeout=10)).send_keys("TSI")
        await (await driver.find_element(By.ID, "txtUserName", timeout=10)).send_keys(username)
        await (await driver.find_element(By.ID, "txtPassword", timeout=10)).send_keys(password)
        await (await driver.find_element(By.ID, "LoginButton", timeout=10)).click()

        log_callback("Verifying credentials...")
        try:
            await driver.find_element(By.XPATH, "//form[@id='form1' or @name='aspnetForm']", timeout=15)
            await asyncio.sleep(2)
        except Exception:
            log_callback("Session lock wait timed out. Attempting fallback dashboard load anyway...")

        # 2. Open DTR Dashboard & Apply Filters
        log_callback("Opening DTR Dashboard...")
        await driver.get("https://my-ices.teligent.com.ph/Portal/DTR.aspx", wait_load=True)

        cutoff_start = await driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtCutOffStart", timeout=10)
        await driver.execute_script("arguments[0].value = arguments[1];", cutoff_start, start_date)

        cutoff_end = await driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtCutOffEnd", timeout=10)
        await driver.execute_script("arguments[0].value = arguments[1];", cutoff_end, end_date)

        await (await driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnViewDTR", timeout=10)).click()
        log_callback("Filters applied. Loading DTR webpage...")
        
        await driver.find_element(By.XPATH, "//table[contains(@id, 'grdBiometricsLog')]", timeout=15)

        excel_pointer = 0
        row_xpath = "//table[contains(@id, 'grdBiometricsLog')]//tr[td and not(contains(@class, 'Header'))]"
        web_rows = await driver.find_elements(By.XPATH, row_xpath)
        total_rows = len(web_rows)
        log_callback(f"DTR Page loaded. Detected {total_rows} calendar days on screen.")

        for web_idx in range(total_rows):
            try:
                current_web_rows = await driver.find_elements(By.XPATH, row_xpath)
                target_web_row = current_web_rows[web_idx]

                try:
                    date_label = await target_web_row.find_element(By.XPATH, ".//span[contains(@id, 'lblLogDate') or contains(@id, 'lblDate')]")
                    row_date_text = await date_label.text
                except Exception:
                    row_date_text = f"Row {web_idx+1}"

                day_label = await target_web_row.find_element(By.XPATH, ".//span[contains(@id, 'lblWorkDay')]")
                day_text = (await day_label.text).strip()
                
                is_weekend = day_text in ["Saturday", "Sunday"]
                is_holiday = False

                log_callback(f"Processing Day {web_idx+1}/{total_rows} -> Portal Date: {row_date_text} ({day_text})...")

                dropdown_selection = None
                time_in_val = "08:30 AM"     
                time_out_val = "05:30 PM"
                remarks_val = ""

                if is_weekend:
                    log_callback("   -> Weekend detected. Automating [Day Off].")
                    dropdown_selection = "Day Off"
                else:
                    if excel_pointer >= len(df):
                        log_callback("   -> ⚠️ No more entries remaining in your Excel file.")
                        continue
                    
                    excel_row = df.iloc[excel_pointer]
                    actual_in = str(excel_row.get('TARGET_TIME_IN', '')).strip()
                    actual_out = str(excel_row.get('TARGET_TIME_OUT', '')).strip()
                    
                    log_callback(f"   -> Consuming Excel Row {excel_pointer+1} (In: {actual_in} | Out: {actual_out})")
                    excel_pointer += 1 

                    if not actual_in or actual_in.lower() == 'nan' or actual_in == "":
                        log_callback("   -> Skipping: Excel cell entry is empty.")
                        continue

                    if "holiday" not in actual_in.lower():
                        try:
                            time_in_val = pd.to_datetime(actual_in).strftime('%I:%M %p')
                        except Exception:
                            time_in_val = actual_in
                        try:
                            if actual_out and actual_out.lower() != 'nan' and actual_out != "":
                                time_out_val = pd.to_datetime(actual_out).strftime('%I:%M %p')
                        except Exception:
                            time_out_val = actual_out

                    is_processed_via_time = False
                    schedule_end_time_str = "17:30:00"
                    
                    if "holiday" not in actual_in.lower():
                        try:
                            parsed_in_time = pd.to_datetime(actual_in).time()
                            time_boundary_830 = pd.to_datetime("08:30:00", format="%H:%M:%S").time()
                            time_boundary_930 = pd.to_datetime("09:30:00", format="%H:%M:%S").time()
                            
                            if parsed_in_time < time_boundary_830:
                                dropdown_selection = "8:30AM-5:30PM"
                                schedule_end_time_str = "17:30:00"
                                is_processed_via_time = True
                            elif parsed_in_time > time_boundary_930:
                                dropdown_selection = "9:30AM-6:30PM"
                                schedule_end_time_str = "18:30:00"
                                is_processed_via_time = True
                        except Exception:
                            pass

                    if not is_processed_via_time:
                        if "holiday" in actual_in.lower():
                            dropdown_selection = "Holiday"
                            is_holiday = True
                        elif "08:30" in actual_in:
                            dropdown_selection = "8:30AM-5:30PM"
                            schedule_end_time_str = "17:30:00"
                        elif "09:00" in actual_in:
                            dropdown_selection = "9:00 AM - 6:00 PM"
                            schedule_end_time_str = "18:00:00"
                        elif "09:30" in actual_in:
                            dropdown_selection = "9:30AM-6:30PM"
                            schedule_end_time_str = "18:30:00"
                        else:
                            dropdown_selection = "8:30AM-5:30PM"
                            schedule_end_time_str = "17:30:00"

                    if not is_holiday and actual_out and actual_out.lower() != 'nan' and actual_out != "":
                        try:
                            actual_in_dt = pd.to_datetime(actual_in)
                            actual_out_dt = pd.to_datetime(actual_out)
                            sched_end_dt = pd.to_datetime(schedule_end_time_str, format="%H:%M:%S")
                            total_elapsed_mins = (actual_out_dt - actual_in_dt).total_seconds() / 60.0
                            if total_elapsed_mins >= 540: 
                                time_delta_mins = (actual_out_dt.hour * 60 + actual_out_dt.minute) - (sched_end_dt.hour * 60 + sched_end_dt.minute)
                                if time_delta_mins > 30:
                                    remarks_val = "overtime"
                        except Exception:
                            pass

                # Action step: click Edit button
                edit_btn = await target_web_row.find_element(By.XPATH, ".//input[@type='image' and @alt='Edit']")
                await edit_btn.click()
                
                dropdown = await driver.find_element(By.XPATH, f"({row_xpath})[{web_idx+1}]//select[contains(@id, 'ddlWorkSchedule')]", timeout=4)
                
                # HIGH SPEED OPTIMIZATION: Match option via JavaScript execution context instantly
                js_select_dropdown = """
                    var select = arguments[0];
                    var textToFind = arguments[1];
                    var updated = false;
                    for (var i = 0; i < select.options.length; i++) {
                        if (select.options[i].text.trim() === textToFind || select.options[i].value === textToFind) {
                            if (select.selectedIndex !== i) {
                                select.selectedIndex = i;
                                updated = true;
                            }
                            break;
                        }
                    }
                    if (updated) {
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                    return updated;
                """
                
                was_changed = await driver.execute_script(js_select_dropdown, dropdown, dropdown_selection)
                
                if was_changed:
                    poll_success = False
                    for _ in range(15):  # Poll every 100ms up to 1.5 seconds maximum
                        await asyncio.sleep(0.1)
                        if not is_weekend and not is_holiday:
                            try:
                                check_el = await driver.find_element(By.XPATH, f"({row_xpath})[{web_idx+1}]//input[contains(@id, 'txtTimeIn')]", timeout=0.1)
                                if check_el:
                                    poll_success = True
                                    break
                            except Exception:
                                pass
                        else:
                            poll_success = True
                            break
                else:
                    await asyncio.sleep(0.05)

                if not is_weekend and not is_holiday:
                    time_in_input = await driver.find_element(By.XPATH, f"({row_xpath})[{web_idx+1}]//input[contains(@id, 'txtTimeIn')]", timeout=3)
                    time_out_input = await driver.find_element(By.XPATH, f"({row_xpath})[{web_idx+1}]//input[contains(@id, 'txtTimeOut')]")
                    remarks_input = await driver.find_element(By.XPATH, f"({row_xpath})[{web_idx+1}]//input[contains(@id, 'txtRemarks')]")

                    await driver.execute_script("arguments[0].value = arguments[1];", time_in_input, time_in_val)
                    await driver.execute_script("arguments[0].value = arguments[1];", time_out_input, time_out_val)
                    await driver.execute_script("arguments[0].value = arguments[1];", remarks_input, remarks_val)
                    
                    log_callback(f"      -> Inputted Values: In [{time_in_val}] | Out [{time_out_val}]")

                update_btn = await driver.find_element(By.XPATH, f"({row_xpath})[{web_idx+1}]//input[@type='image' and @alt='Update']")
                await update_btn.click()
                
                await driver.find_element(By.XPATH, f"({row_xpath})[{web_idx+1}]//input[@type='image' and @alt='Edit']", timeout=5)
                log_callback(f"   -> Success: DTR Row saved.")

            except Exception as row_ex:
                log_callback(f"   -> ERROR at row position {web_idx+1}: {str(row_ex)}")
                await asyncio.sleep(0.5)
                continue

        log_callback("Finished running automation. DTR complete!")
        gui_instance.root.after(0, lambda: messagebox.showinfo("Automation Success", "Finished running DTR automation!\nAll entries have been uploaded."))

# -----------------------------------------------------------------------------
# GUI Core Dashboard Panels
# -----------------------------------------------------------------------------
class DtrAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DTR Automation")
        self.root.geometry("550x520")
        self.root.resizable(False, False)
        
        style = ttk.Style()
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Action.TButton", font=("Segoe UI", 10, "bold"), foreground="green")
        
        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Username:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.username_entry = ttk.Entry(main_frame, width=32)
        self.username_entry.insert(0, "200000436")
        self.username_entry.grid(row=0, column=1, sticky=tk.E, pady=4)

        ttk.Label(main_frame, text="Password:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.password_entry = ttk.Entry(main_frame, show="*", width=32)
        self.password_entry.insert(0, "12#manila")
        self.password_entry.grid(row=1, column=1, sticky=tk.E, pady=4)

        ttk.Separator(main_frame, orient='horizontal').grid(row=2, column=0, columnspan=2, sticky="ew", pady=8)

        ttk.Label(main_frame, text="Cutoff Start Date (MM/DD/YYYY):").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.start_date_entry = ttk.Entry(main_frame, width=32)
        self.start_date_entry.insert(0, "06/16/2026")
        self.start_date_entry.grid(row=3, column=1, sticky=tk.E, pady=4)

        ttk.Label(main_frame, text="Cutoff End Date (MM/DD/YYYY):").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.end_date_entry = ttk.Entry(main_frame, width=32)
        self.end_date_entry.insert(0, "06/30/2026")
        self.end_date_entry.grid(row=4, column=1, sticky=tk.E, pady=4)

        ttk.Separator(main_frame, orient='horizontal').grid(row=5, column=0, columnspan=2, sticky="ew", pady=8)

        ttk.Label(main_frame, text="Target Excel File:").grid(row=6, column=0, sticky=tk.W, pady=4)
        file_picker_frame = ttk.Frame(main_frame)
        file_picker_frame.grid(row=6, column=1, sticky=tk.E, pady=4)
        
        self.file_path_entry = ttk.Entry(file_picker_frame, width=22)
        self.file_path_entry.pack(side=tk.LEFT, padx=(0, 4))
        
        self.browse_btn = ttk.Button(file_picker_frame, text="Browse...", width=8, command=self.open_file_selector)
        self.browse_btn.pack(side=tk.LEFT)

        self.run_button = ttk.Button(main_frame, text="RUN FILE AUTOMATION", style="Action.TButton", command=self.dispatch_worker_thread)
        self.run_button.grid(row=7, column=0, columnspan=2, pady=15, sticky="ew")

        ttk.Label(main_frame, text="Realtime Activity logs:").grid(row=8, column=0, sticky=tk.W)
        self.log_text = tk.Text(main_frame, height=10, width=70, state=tk.DISABLED, bg="#F4F6F7", font=("Consolas", 9))
        self.log_text.grid(row=9, column=0, columnspan=2, pady=4)

    def open_file_selector(self):
        selected_file = filedialog.askopenfilename(
            title="Select DTR File",
            filetypes=[("Excel Workspace files", "*.xlsx *.xls"), ("CSV sheets", "*.csv")]
        )
        if selected_file:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, selected_file)

    def write_log(self, text):
        self.root.after(0, lambda: self._unsafe_write(text))

    def _unsafe_write(self, text):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def dispatch_worker_thread(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        start_date = self.start_date_entry.get().strip()
        end_date = self.end_date_entry.get().strip()
        file_path = self.file_path_entry.get().strip()

        if not all([username, password, start_date, end_date, file_path]):
            messagebox.showwarning("Missing Fields", "Please make sure to fill out all the fields.")
            return

        self.run_button.config(state=tk.DISABLED)
        
        def thread_wrapper():
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(
                    run_automation(self, username, password, start_date, end_date, file_path, self.write_log)
                )
            except Exception as loop_ex:
                self.write_log(f"Fatal Thread Execution Error: {str(loop_ex)}")
                self.root.after(0, lambda: messagebox.showerror("Automation Error", f"A fatal crash occurred:\n{str(loop_ex)}"))
            finally:
                self.root.after(0, lambda: self.run_button.config(state=tk.NORMAL))

        threading.Thread(target=thread_wrapper, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = DtrAutomationGUI(root)
    root.mainloop()