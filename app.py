import sys
import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from ldap3 import Server, Connection, ALL, NTLM

class ADSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Active Directory Targeted Search - SONATRACH")
        self.root.geometry("800x620")
        
        style = ttk.Style()
        style.theme_use("vista" if sys.platform == "win32" else "clam")
        
        # --- Connection Frame ---
        conn_frame = ttk.LabelFrame(root, text=" Server Connection Configuration ", padding=10)
        conn_frame.pack(fill="x", padx=15, pady=10)
        
        ttk.Label(conn_frame, text="LDAP Server/Domain:").grid(row=0, column=0, sticky="w", pady=2)
        self.server_entry = ttk.Entry(conn_frame, width=30)
        self.server_entry.grid(row=0, column=1, padx=5, pady=2)
        self.server_entry.insert(0, "corp.sonatrach.dz")

        ttk.Label(conn_frame, text="Base DN:").grid(row=0, column=2, sticky="w", pady=2)
        self.base_entry = ttk.Entry(conn_frame, width=30)
        self.base_entry.grid(row=0, column=3, padx=5, pady=2)
        self.base_entry.insert(0, "DC=corp,DC=sonatrach,DC=dz")

        ttk.Label(conn_frame, text="Domain User:").grid(row=1, column=0, sticky="w", pady=2)
        self.user_entry = ttk.Entry(conn_frame, width=30)
        self.user_entry.grid(row=1, column=1, padx=5, pady=2)
        self.user_entry.insert(0, "SONATRACH\\f.berraoui")

        ttk.Label(conn_frame, text="Windows Password:").grid(row=1, column=2, sticky="w", pady=2)
        self.pass_entry = ttk.Entry(conn_frame, show="*", width=30)
        self.pass_entry.grid(row=1, column=3, padx=5, pady=2)

        # --- Filter Frame ---
        filter_frame = ttk.LabelFrame(root, text=" Target Search Criteria ", padding=10)
        filter_frame.pack(fill="x", padx=15, pady=5)

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
        
        server_target = self.server_entry.get().strip()
        base_dn = self.base_entry.get().strip()
        user_identity = self.user_entry.get().strip()
        user_password = self.pass_entry.get()
        
        search_user = self.username_entry.get().strip()
        search_mat = self.matricule_entry.get().strip()
        search_name = self.name_entry.get().strip()

        if not user_password:
            messagebox.showwarning("Password Required", "Please enter your domain password to authenticate.")
            return

        if not search_user and not search_mat and not search_name:
            messagebox.showwarning("Input Required", "Please enter a Username, Matricule, or Name to filter.")
            return

        ldap_filter = "(&(objectCategory=person)(objectClass=user)"
        if search_user:
            ldap_filter += f"(sAMAccountName=*{search_user}*)"
        if search_mat:
            ldap_filter += f"(employeeID=*{search_mat}*)"
        if search_name:
            ldap_filter += f"(displayName=*{search_name}*)"
        ldap_filter += ")"

        # Cycle ports safely if one is firewalled
        ports_to_try = [389, 3268, 636]
        connection_success = False
        last_error = ""

        for port in ports_to_try:
            try:
                use_ssl = True if port == 636 else False
                server = Server(server_target, port=port, use_ssl=use_ssl, get_info=ALL)
                
                # NTLM explicit authentication bind
                with Connection(server, user=user_identity, password=user_password, authentication=NTLM, auto_bind=True) as conn:
                    attributes = ['sAMAccountName', 'employeeID', 'displayName', 'department', 'mail']
                    conn.search(search_base=base_dn, search_filter=ldap_filter, attributes=attributes)
                    
                    for entry in conn.entries:
                        row_values = (
                            str(entry.sAMAccountName or ''),
                            str(entry.employeeID or ''),
                            str(entry.displayName or ''),
                            str(entry.department or ''),
                            str(entry.mail or '')
                        )
                        self.tree.insert("", "end", values=row_values)
                        self.results_data.append(row_values)
                    
                    connection_success = True
                    break 
            except Exception as e:
                last_error = str(e)
                continue

        if connection_success:
            if self.results_data:
                self.export_btn.config(state="normal")
            else:
                self.export_btn.config(state="disabled")
                messagebox.showwarning("No Matches", "Connection succeeded, but no records matched your filter criteria.")
        else:
            messagebox.showerror("Connection Failed", f"Could not connect to directory service:\n{last_error}")

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
