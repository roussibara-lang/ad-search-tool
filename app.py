import sys
import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Try importing the native Windows COM client
try:
    import win32com.client
except ImportError:
    messagebox.showerror(
        "Missing Dependency", 
        "The 'pywin32' library is missing from this build environment."
    )
    sys.exit(1)

class ADSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Active Directory Admin Search - SONATRACH")
        self.root.geometry("800x550")
        
        style = ttk.Style()
        style.theme_use("vista" if sys.platform == "win32" else "clam")
        
        # --- Filter Frame ---
        filter_frame = ttk.LabelFrame(root, text=" Targeted Search Fields (Password-Free Admin Mode) ", padding=10)
        filter_frame.pack(fill="x", padx=15, pady=15)

        ttk.Label(filter_frame, text="Username (sAMAccountName):").grid(row=0, column=0, sticky="w", pady=2)
        self.username_entry = ttk.Entry(filter_frame, width=25)
        self.username_entry.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(filter_frame, text="Matricule (Employee ID):").grid(row=0, column=2, sticky="w", pady=2)
        self.matricule_entry = ttk.Entry(filter_frame, width=25)
        self.matricule_entry.grid(row=0, column=3, padx=5, pady=2)

        ttk.Label(filter_frame, text="Display Name / Full Name:").grid(row=1, column=0, sticky="w", pady=2)
        self.name_entry = ttk.Entry(filter_frame, width=25)
        self.name_entry.grid(row=1, column=1, padx=5, pady=2)

        # --- Actions ---
        btn_frame = ttk.Frame(root, padding=5)
        btn_frame.pack(fill="x", padx=15, pady=5)
        
        self.search_btn = ttk.Button(btn_frame, text="Search User", command=self.search_ad)
        self.search_btn.pack(side="left", padx=5)
        
        self.export_btn = ttk.Button(btn_frame, text="Export CSV", command=self.export_csv, state="disabled")
        self.export_btn.pack(side="left", padx=5)

        # --- Results Table ---
        table_frame = ttk.Frame(root, padding=10)
        table_frame.pack(fill="both", expand=True, padx=5)

        columns = ("username", "matricule", "fullname", "department", "email")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        
        self.tree.heading("username", text="Username")
        self.tree.heading("matricule", text="Matricule")
        self.tree.heading("fullname", text="Full Name")
        self.tree.heading("department", text="Department")
        self.tree.heading("email", text="Email")
        
        self.tree.column("username", width=120)
        self.tree.column("matricule", width=100)
        self.tree.column("fullname", width=160)
        self.tree.column("department", width=140)
        self.tree.column("email", width=200)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.results_data = []

    def search_ad(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.results_data.clear()
        
        search_user = self.username_entry.get().strip()
        search_mat = self.matricule_entry.get().strip()
        search_name = self.name_entry.get().strip()

        if not search_user and not search_mat and not search_name:
            messagebox.showwarning("Input Required", "Please enter a Username, Matricule, or Name to filter.")
            return

        # Build LDAP query targeting corp.sonatrach.dz
        ldap_filter = "(&(objectCategory=person)(objectClass=user)"
        if search_user:
            ldap_filter += f"(sAMAccountName=*{search_user}*)"
        if search_mat:
            ldap_filter += f"(employeeID=*{search_mat}*)"
        if search_name:
            ldap_filter += f"(displayName=*{search_name}*)"
        ldap_filter += ")"

        try:
            # Connect natively using Active Directory's system COM provider
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Provider = "ADSDSOObject"
            conn.Open("Active Directory Provider")
            
            cmd = win32com.client.Dispatch("ADODB.Command")
            cmd.ActiveConnection = conn
            
            # Formulate the query string using the live RootDSE
            query = f"<LDAP://corp.sonatrach.dz/DC=corp,DC=sonatrach,DC=dz>;{ldap_filter};sAMAccountName,employeeID,displayName,department,mail;subtree"
            cmd.CommandText = query
            
            recordset, _ = cmd.Execute()
            
            while not recordset.EOF:
                # Safely pull AD attribute values
                user_id = recordset.Fields("sAMAccountName").Value
                matricule = recordset.Fields("employeeID").Value
                fullname = recordset.Fields("displayName").Value
                dept = recordset.Fields("department").Value
                email = recordset.Fields("mail").Value

                row_values = (
                    str(user_id or ''),
                    str(matricule or ''),
                    str(fullname or ''),
                    str(dept or ''),
                    str(email or '')
                )
                self.tree.insert("", "end", values=row_values)
                self.results_data.append(row_values)
                recordset.MoveNext()

            recordset.Close()
            conn.Close()

            if self.results_data:
                self.export_btn.config(state="normal")
            else:
                self.export_btn.config(state="disabled")
                messagebox.showwarning("No Matches", "No matching domain accounts found.")

        except Exception as e:
            messagebox.showerror("Execution Failed", f"Could not perform lookup operation natively:\n{str(e)}")

    def export_csv(self):
        if not self.results_data:
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV File", "*.csv")])
        if file_path:
            with open(file_path, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Username", "Matricule", "Full Name", "Department", "Email"])
                writer.writerows(self.results_data)
            messagebox.showinfo("Exported", "Data target exported successfully.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ADSearchApp(root)
    root.mainloop()
