import sys
import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import win32com.client
    # Python uses this to interact with Windows Active Directory Service Interfaces (ADSI)
except ImportError:
    messagebox.showerror(
        "Missing Dependency", 
        "The 'pywin32' library is missing from this build environment."
    )
    sys.exit(1)

class ADSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Active Directory Admin Panel - SONATRACH")
        self.root.geometry("900x650")
        self.root.minsize(800, 600)
        
        style = ttk.Style()
        style.theme_use("vista" if sys.platform == "win32" else "clam")
        
        # --- Filter Frame ---
        filter_frame = ttk.LabelFrame(root, text=" Targeted Search Fields (Password-Free Admin Mode) ", padding=10)
        filter_frame.pack(fill="x", padx=15, pady=10)

        ttk.Label(filter_frame, text="Username (sAMAccountName):").grid(row=0, column=0, sticky="w", pady=2)
        self.username_entry = ttk.Entry(filter_frame, width=25)
        self.username_entry.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(filter_frame, text="Matricule (Employee ID):").grid(row=0, column=2, sticky="w", pady=2)
        self.matricule_entry = ttk.Entry(filter_frame, width=25)
        self.matricule_entry.grid(row=0, column=3, padx=5, pady=2)

        ttk.Label(filter_frame, text="Display Name / Full Name:").grid(row=1, column=0, sticky="w", pady=2)
        self.name_entry = ttk.Entry(filter_frame, width=25)
        self.name_entry.grid(row=1, column=1, padx=5, pady=2)

        # --- Actions Frame ---
        btn_frame = ttk.Frame(root, padding=5)
        btn_frame.pack(fill="x", padx=15, pady=5)
        
        self.search_btn = ttk.Button(btn_frame, text="Search User", command=self.search_ad)
        self.search_btn.pack(side="left", padx=5)
        
        # New administrative action buttons
        self.prop_btn = ttk.Button(btn_frame, text="Show Properties", command=self.show_properties, state="disabled")
        self.prop_btn.pack(side="left", padx=5)

        self.reset_btn = ttk.Button(btn_frame, text="Reset Password", command=self.reset_password, state="disabled")
        self.reset_btn.pack(side="left", padx=5)

        self.export_btn = ttk.Button(btn_frame, text="Export CSV", command=self.export_csv, state="disabled")
        self.export_btn.pack(side="left", padx=5)

        # --- Results Table ---
        table_frame = ttk.Frame(root, padding=10)
        table_frame.pack(fill="both", expand=True, padx=5)

        # We keep Distinguished Name hidden in the UI but saved in internal state for targeting operations
        columns = ("username", "matricule", "fullname", "department", "email")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("username", text="Username")
        self.tree.heading("matricule", text="Matricule")
        self.tree.heading("fullname", text="Full Name")
        self.tree.heading("department", text="Department")
        self.tree.heading("email", text="Email")
        
        self.tree.column("username", width=120)
        self.tree.column("matricule", width=100)
        self.tree.column("fullname", width=180)
        self.tree.column("department", width=140)
        self.tree.column("email", width=220)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Event triggers
        self.tree.bind("<<TreeviewSelect>>", self.on_user_select)
        self.tree.bind("<Double-1>", lambda event: self.show_properties())

        self.results_data = []
        self.selected_dn = None
        self.selected_username = None

    def on_user_select(self, event):
        """Enables and disables action buttons depending on selected row state."""
        selected_item = self.tree.selection()
        if selected_item:
            item_data = self.tree.item(selected_item[0])
            username = item_data['values'][0]
            
            # Locate original record to grab the distinguishedName (DN)
            for row in self.results_data:
                if row[0] == username:
                    self.selected_dn = row[5] # Index 5 holds the DN
                    self.selected_username = username
                    self.prop_btn.config(state="normal")
                    self.reset_btn.config(state="normal")
                    return
        else:
            self.selected_dn = None
            self.selected_username = None
            self.prop_btn.config(state="disabled")
            self.reset_btn.config(state="disabled")

    def search_ad(self):
        """Queries corp.sonatrach.dz natively for target records."""
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.results_data.clear()
        self.selected_dn = None
        self.selected_username = None
        self.prop_btn.config(state="disabled")
        self.reset_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        
        search_user = self.username_entry.get().strip()
        search_mat = self.matricule_entry.get().strip()
        search_name = self.name_entry.get().strip()

        if not search_user and not search_mat and not search_name:
            messagebox.showwarning("Input Required", "Please enter a Username, Matricule, or Name to filter.")
            return

        # Build clean Active Directory search syntax
        ldap_filter = "(&(objectCategory=person)(objectClass=user)"
        if search_user:
            ldap_filter += f"(sAMAccountName=*{search_user}*)"
        if search_mat:
            ldap_filter += f"(employeeID=*{search_mat}*)"
        if search_name:
            ldap_filter += f"(displayName=*{search_name}*)"
        ldap_filter += ")"

        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Provider = "ADSDSOObject"
            conn.Open("Active Directory Provider")
            
            cmd = win32com.client.Dispatch("ADODB.Command")
            cmd.ActiveConnection = conn
            
            # Querying the DN along with target visual parameters
            query = f"<LDAP://corp.sonatrach.dz/DC=corp,DC=sonatrach,DC=dz>;{ldap_filter};sAMAccountName,employeeID,displayName,department,mail,distinguishedName;subtree"
            cmd.CommandText = query
            
            recordset, _ = cmd.Execute()
            
            while not recordset.EOF:
                user_id = recordset.Fields("sAMAccountName").Value
                matricule = recordset.Fields("employeeID").Value
                fullname = recordset.Fields("displayName").Value
                dept = recordset.Fields("department").Value
                email = recordset.Fields("mail").Value
                dn = recordset.Fields("distinguishedName").Value

                row_values_ui = (
                    str(user_id or ''),
                    str(matricule or ''),
                    str(fullname or ''),
                    str(dept or ''),
                    str(email or '')
                )
                
                self.tree.insert("", "end", values=row_values_ui)
                # Keep the full list including Distinguished Name (DN) internally
                self.results_data.append((*row_values_ui, str(dn or '')))
                recordset.MoveNext()

            recordset.Close()
            conn.Close()

            if self.results_data:
                self.export_btn.config(state="normal")
            else:
                messagebox.showwarning("No Matches", "No matching domain accounts found.")

        except Exception as e:
            messagebox.showerror("Execution Failed", f"Could not perform lookup operation natively:\n{str(e)}")

    def show_properties(self):
        """Binds to the selected user's LDAP distinguishedName to pull extensive properties."""
        if not self.selected_dn:
            return
            
        try:
            # Bind directly to the user object using active Windows security context
            user_obj = win32com.client.GetObject(f"LDAP://{self.selected_dn}")
            
            # Generate properties popup window
            prop_win = tk.Toplevel(self.root)
            prop_win.title(f"User Properties - {self.selected_username}")
            prop_win.geometry("450x450")
            prop_win.resizable(False, False)
            prop_win.grab_set() # Lock focus to this window
            
            main_frame = ttk.Frame(prop_win, padding=20)
            main_frame.pack(fill="both", expand=True)

            ttk.Label(main_frame, text="Active Directory Object Details", font=("Segoe UI", 12, "bold")).pack(pady=(0, 15))

            details_frame = ttk.LabelFrame(main_frame, text=" Account Properties ", padding=15)
            details_frame.pack(fill="both", expand=True)

            # Safely fetch properties using getattr or string values
            def get_ad_prop(prop_name):
                try:
                    return str(user_obj.Get(prop_name) or 'N/A')
                except Exception:
                    return 'N/A'

            properties = [
                ("Full Name:", "displayName"),
                ("Username:", "sAMAccountName"),
                ("Matricule:", "employeeID"),
                ("Email:", "mail"),
                ("Department:", "department"),
                ("Title:", "title"),
                ("Phone:", "telephoneNumber")
            ]

            # Parse Account Status (Enabled / Disabled)
            account_status = "Enabled"
            try:
                uac = int(user_obj.Get("userAccountControl"))
                if uac & 2: # ADS_UF_ACCOUNTDISABLE flag is decimal 2
                    account_status = "Disabled"
            except Exception:
                account_status = "Unknown"

            # Render properties cleanly in grid
            row_idx = 0
            for label, ldap_attr in properties:
                ttk.Label(details_frame, text=label, font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=0, sticky="w", pady=4, padx=5)
                ttk.Label(details_frame, text=get_ad_prop(ldap_attr), wraplength=250).grid(row=row_idx, column=1, sticky="w", pady=4, padx=5)
                row_idx += 1

            # Account Status Grid Row
            ttk.Label(details_frame, text="Account Status:", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=0, sticky="w", pady=4, padx=5)
            status_lbl = ttk.Label(details_frame, text=account_status, foreground="green" if account_status == "Enabled" else "red", font=("Segoe UI", 9, "bold"))
            status_lbl.grid(row=row_idx, column=1, sticky="w", pady=4, padx=5)

            # Close Button
            ttk.Button(main_frame, text="Close Properties", command=prop_win.destroy).pack(pady=(15, 0))

        except Exception as e:
            messagebox.showerror("Error Opening Properties", f"Could not bind to user object:\n{str(e)}")

    def reset_password(self):
        """Binds natively to change selected AD account passwords."""
        if not self.selected_dn:
            return

        # Create Password Reset pop-up window
        reset_win = tk.Toplevel(self.root)
        reset_win.title(f"Reset Password - {self.selected_username}")
        reset_win.geometry("400x250")
        reset_win.resizable(False, False)
        reset_win.grab_set()

        main_frame = ttk.Frame(reset_win, padding=20)
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text=f"Set New Password for {self.selected_username}", font=("Segoe UI", 10, "bold")).pack(pady=(0, 10))

        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill="x", pady=5)

        ttk.Label(input_frame, text="New Password:").grid(row=0, column=0, sticky="w", pady=5)
        new_pass_entry = ttk.Entry(input_frame, show="*", width=25)
        new_pass_entry.grid(row=0, column=1, padx=10, pady=5)
        new_pass_entry.focus()

        ttk.Label(input_frame, text="Confirm Password:").grid(row=1, column=0, sticky="w", pady=5)
        conf_pass_entry = ttk.Entry(input_frame, show="*", width=25)
        conf_pass_entry.grid(row=1, column=1, padx=10, pady=5)

        # Force password change flag checkbox
        force_change_var = tk.BooleanVar(value=True)
        force_change_cb = ttk.Checkbutton(main_frame, text="User must change password at next logon", variable=force_change_var)
        force_change_cb.pack(pady=10)

        def submit_reset():
            new_pwd = new_pass_entry.get()
            conf_pwd = conf_pass_entry.get()

            if not new_pwd:
                messagebox.showerror("Validation Error", "Password field cannot be empty.")
                return
            if new_pwd != conf_pwd:
                messagebox.showerror("Validation Error", "Passwords do not match. Please verify typing.")
                return

            try:
                # Bind directly to AD User Object via the Distinguished Name (DN)
                user_obj = win32com.client.GetObject(f"LDAP://{self.selected_dn}")
                
                # Active Directory's native password modification method
                user_obj.SetPassword(new_pwd)
                
                # Apply forced password change flags if checked
                if force_change_var.get():
                    # Writing integer value '0' directly to the 'pwdLastSet' attribute mandates reset on next auth
                    user_obj.Put("pwdLastSet", 0)
                else:
                    # Write '-1' to keep current state intact
                    user_obj.Put("pwdLastSet", -1)
                
                # Save changes back to Active Directory schema
                user_obj.SetInfo()

                messagebox.showinfo("Success", f"Password for {self.selected_username} has been reset successfully.")
                reset_win.destroy()

            except Exception as e:
                messagebox.showerror("AD Reset Failed", f"Active Directory refused password change request:\n{str(e)}")

        # Buttons Frame
        btn_box = ttk.Frame(main_frame)
        btn_box.pack(pady=10)
        
        ttk.Button(btn_box, text="Change Password", command=submit_reset).pack(side="left", padx=5)
        ttk.Button(btn_box, text="Cancel", command=reset_win.destroy).pack(side="left", padx=5)

    def export_csv(self):
        if not self.results_data:
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV File", "*.csv")])
        if file_path:
            with open(file_path, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Username", "Matricule", "Full Name", "Department", "Email"])
                # We do not export Distinguished Name to the customer-facing CSV
                for row in self.results_data:
                    writer.writerow(row[:5])
            messagebox.showinfo("Exported", "Data target exported successfully.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ADSearchApp(root)
    root.mainloop()
