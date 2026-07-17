import sys
import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Attempt to handle dependencies gracefully
try:
    from ldap3 import Server, Connection, ALL, SAFEST_AUTH
except ImportError:
    messagebox.showerror(
        "Missing Dependency", 
        "The 'ldap3' library is missing.\nPlease run: pip install ldap3"
    )
    sys.exit(1)

class ADSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Active Directory User Lookup Tool")
        self.root.geometry("750x580")
        
        # --- Style Configuration ---
        style = ttk.Style()
        style.theme_use("vista" if sys.platform == "win32" else "clam")
        
        # --- Connection Frame ---
        conn_frame = ttk.LabelFrame(root, text=" Server Connection Details ", padding=10)
        conn_frame.pack(fill="x", padx=15, pady=10)
        
        ttk.Label(conn_frame, text="LDAP Server:").grid(row=0, column=0, sticky="w", pady=2)
        self.server_entry = ttk.Entry(conn_frame, width=28)
        self.server_entry.grid(row=0, column=1, padx=5, pady=2)
        self.server_entry.insert(0, "ldap://192.168.1.1") # Standard internal IP placeholder

        ttk.Label(conn_frame, text="Base DN:").grid(row=0, column=2, sticky="w", pady=2)
        self.base_entry = ttk.Entry(conn_frame, width=28)
        self.base_entry.grid(row=0, column=3, padx=5, pady=2)
        self.base_entry.insert(0, "DC=company,DC=local")

        ttk.Label(conn_frame, text="Domain User:").grid(row=1, column=0, sticky="w", pady=2)
        self.user_entry = ttk.Entry(conn_frame, width=28)
        self.user_entry.grid(row=1, column=1, padx=5, pady=2)
        self.user_entry.insert(0, "DOMAIN\\Administrator")

        ttk.Label(conn_frame, text="Password:").grid(row=1, column=2, sticky="w", pady=2)
        self.pass_entry = ttk.Entry(conn_frame, show="*", width=28)
        self.pass_entry.grid(row=1, column=3, padx=5, pady=2)

        # --- Filter Frame ---
        filter_frame = ttk.LabelFrame(root, text=" Filter Criteria ", padding=10)
        filter_frame.pack(fill="x", padx=15, pady=5)

        ttk.Label(filter_frame, text="Department:").grid(row=0, column=0, sticky="w", pady=2)
        self.dept_entry = ttk.Entry(filter_frame, width=28)
        self.dept_entry.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(filter_frame, text="State/Province:").grid(row=0, column=2, sticky="w", pady=2)
        self.state_entry = ttk.Entry(filter_frame, width=28)
        self.state_entry.grid(row=0, column=3, padx=5, pady=2)

        # --- Actions ---
        btn_frame = ttk.Frame(root, padding=5)
        btn_frame.pack(fill="x", padx=15, pady=5)
        
        self.search_btn = ttk.Button(btn_frame, text="Fetch AD Users", command=self.search_ad)
        self.search_btn.pack(side="left", padx=5)
        
        self.export_btn = ttk.Button(btn_frame, text="Export CSV Audience", command=self.export_csv, state="disabled")
        self.export_btn.pack(side="left", padx=5)

        # --- Results Table ---
        table_frame = ttk.Frame(root, padding=10)
        table_frame.pack(fill="both", expand=True, padx=5)

        columns = ("username", "firstname", "lastname", "department", "state", "email")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        
        self.tree.heading("username", text="Username")
        self.tree.heading("firstname", text="First Name")
        self.tree.heading("lastname", text="Last Name")
        self.tree.heading("department", text="Department")
        self.tree.heading("state", text="State")
        self.tree.heading("email", text="Email")
        
        self.tree.column("username", width=100)
        self.tree.column("firstname", width=100)
        self.tree.column("lastname", width=100)
        self.tree.column("department", width=120)
        self.tree.column("state", width=80)
        self.tree.column("email", width=180)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.results_data = []

    def search_ad(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.results_data.clear()
        
        server_url = self.server_entry.get()
        base_dn = self.base_entry.get()
        username = self.user_entry.get()
        password = self.pass_entry.get()
        
        ldap_filter = "(&(objectCategory=person)(objectClass=user)"
        if self.dept_entry.get():
            ldap_filter += f"(department={self.dept_entry.get()})"
        if self.state_entry.get():
            ldap_filter += f"(st={self.state_entry.get()})"
        ldap_filter += ")"

        try:
            server = Server(server_url, get_info=ALL)
            with Connection(server, user=username, password=password, authentication=SAFEST_AUTH, auto_bind=True) as conn:
                attributes = ['sAMAccountName', 'givenName', 'sn', 'department', 'st', 'mail']
                conn.search(search_base=base_dn, search_filter=ldap_filter, attributes=attributes)
                
                for entry in conn.entries:
                    row_values = (
                        str(entry.sAMAccountName or ''),
                        str(entry.givenName or ''),
                        str(entry.sn or ''),
                        str(entry.department or ''),
                        str(entry.st or ''),
                        str(entry.mail or '')
                    )
                    self.tree.insert("", "end", values=row_values)
                    self.results_data.append(row_values)
                
                if self.results_data:
                    self.export_btn.config(state="normal")
                    messagebox.showinfo("Done", f"Pulled {len(self.results_data)} matching users successfully.")
                else:
                    self.export_btn.config(state="disabled")
                    messagebox.showwarning("Empty", "No Active Directory entries matched your filter parameters.")
                    
        except Exception as e:
            messagebox.showerror("AD Connection Error", f"Could not pull data from the server:\n{str(e)}")

    def export_csv(self):
        if not self.results_data:
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV File", "*.csv")])
        if file_path:
            with open(file_path, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Username", "First Name", "Last Name", "Department", "State/Province", "Email"])
                writer.writerows(self.results_data)
            messagebox.showinfo("Exported", "Your targeted audience file has been saved.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ADSearchApp(root)
    root.mainloop()