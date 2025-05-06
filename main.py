import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import requests
import json
import os
import tksheet  # pip install tksheet

CONFIG_FILE = "config.json"

class MondayUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Monday.com Uploader")
        self.api_key = ""
        self.users = {}  # Store user mapping
        self.user_emails = {}  # Store email to user mapping

        self.load_config()

        # API Key
        tk.Label(root, text="Monday.com API Key").pack()
        self.api_key_entry = tk.Entry(root, width=50)
        self.api_key_entry.insert(0, self.api_key)
        self.api_key_entry.pack()

        # Load Boards Button
        tk.Button(root, text="Load Boards", command=self.load_boards).pack(pady=5)

        # Board dropdown
        self.board_var = tk.StringVar()
        self.board_dropdown = tk.OptionMenu(root, self.board_var, "Load boards first...")
        self.board_dropdown.pack()

        # Select file to upload
        tk.Button(root, text="Select Excel/CSV File", command=self.load_file).pack(pady=5)

        # Open spreadsheet editor
        tk.Button(root, text="Open Spreadsheet Editor", command=self.open_table_editor).pack(pady=10)

        self.status_label = tk.Label(root, text="", fg="blue")
        self.status_label.pack()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                self.api_key = data.get("api_key", "")
                self.last_board_id = data.get("last_board_id", None)
        else:
            self.api_key = ""
            self.last_board_id = None

    def save_config(self):
        data = {
            "api_key": self.api_key_entry.get().strip(),
            "last_board_id": self.board_var.get()
        }
        with open(CONFIG_FILE, "w") as f:
            json.dump(data, f)

    def fetch_users(self):
        """Fetch users from Monday.com and create mappings"""
        api_key = self.api_key_entry.get().strip()
        if not api_key:
            return

        query = """
        {
          users {
            id
            name
            email
          }
        }
        """

        headers = {
            "Authorization": api_key,
            "Content-Type": "application/json"
        }

        try:
            response = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query})
            data = response.json()

            if "errors" in data:
                print("Error fetching users:", data["errors"])
                return

            self.users = {}
            self.user_emails = {}
            
            for user in data["data"]["users"]:
                self.users[user["name"].lower()] = user["id"]
                if user["email"]:
                    self.user_emails[user["email"].lower()] = user["id"]
            
            print(f"Loaded {len(self.users)} users")
            print("Available users:", list(self.users.keys()))
        except Exception as e:
            print(f"Error fetching users: {e}")

    def load_boards(self):
        api_key = self.api_key_entry.get().strip()
        if not api_key:
            messagebox.showwarning("Missing API Key", "Enter your API key before loading boards.")
            return

        # Fetch users first
        self.fetch_users()

        query = """
        {
          boards {
            id
            name
          }
        }
        """

        headers = {
            "Authorization": api_key,
            "Content-Type": "application/json"
        }

        response = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query})
        data = response.json()

        if "errors" in data:
            messagebox.showerror("API Error", str(data["errors"]))
            return

        boards = data["data"]["boards"]
        menu = self.board_dropdown["menu"]
        menu.delete(0, "end")

        for board in boards:
            menu.add_command(label=f"{board['name']} ({board['id']})",
                             command=lambda value=board['id']: self.board_var.set(value))

        if boards:
            self.board_var.set(str(boards[0]['id']))

    def get_user_id(self, value):
        """Get user ID from name or email"""
        if pd.isna(value) or value == "":
            return None
            
        value = str(value).lower().strip()
        
        # Try to find by name
        if value in self.users:
            return self.users[value]
            
        # Try to find by email
        if value in self.user_emails:
            return self.user_emails[value]
            
        # Try to find by partial name match
        for name, user_id in self.users.items():
            if value in name or name in value:
                return user_id
                
        return None

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")])
        if not file_path:
            return

        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            self.upload_to_monday(df)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
            return

    def upload_to_monday(self, df):
        self.save_config()

        api_key = self.api_key_entry.get().strip()
        board_id = self.board_var.get()
        if not api_key or not board_id:
            messagebox.showwarning("Missing Info", "Please enter API key and select a board.")
            return

        # Ensure we have user data
        if not self.users:
            self.fetch_users()

        headers = {
            "Authorization": api_key,
            "Content-Type": "application/json"
        }

        # Fetch columns from board
        query = """
        query ($boardId: [ID!]) {
          boards(ids: $boardId) {
            columns {
              id
              title
              type
            }
          }
        }
        """
        variables = {"boardId": [str(board_id)]}
        response = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query, "variables": variables})
        result = response.json()

        if "errors" in result:
            messagebox.showerror("API Error", str(result["errors"]))
            return

        columns = result["data"]["boards"][0]["columns"]
        column_map = {col["title"]: col for col in columns}

        url = "https://api.monday.com/v2"
        success_count = 0

        for _, row in df.iterrows():
            item_name = str(row.get('Name', 'Untitled Item'))
            column_values = {}

            for col in df.columns:
                val = row[col]
                if pd.isna(val) or col not in column_map:
                    continue

                monday_col = column_map[col]
                col_id = monday_col["id"]
                col_type = monday_col["type"]

                try:
                    if col_type == "board-relation":
                        item_ids = [int(i.strip()) for i in str(val).split(",") if i.strip().isdigit()]
                        column_values[col_id] = json.dumps({"item_ids": item_ids})
                    elif col_type == "date":
                        if isinstance(val, pd.Timestamp):
                            val = val.strftime('%Y-%m-%d')
                        column_values[col_id] = val
                    elif col_type == "number":
                        column_values[col_id] = float(val)
                    elif col_type == "status":
                        column_values[col_id] = json.dumps({"label": str(val)})
                    elif col_type == "people":
                        # Handle people column type
                        if pd.isna(val) or val == "":
                            column_values[col_id] = json.dumps({"personsAndTeams": []})
                        else:
                            # Handle multiple people (comma-separated)
                            people_ids = []
                            for person in str(val).split(','):
                                user_id = self.get_user_id(person.strip())
                                if user_id:
                                    people_ids.append({"id": int(user_id), "kind": "person"})
                            
                            column_values[col_id] = json.dumps({"personsAndTeams": people_ids})
                    elif col_type == "email":
                        # Email needs to be sent as a JSON object with email and text
                        email = str(val).strip()
                        if not email:  # Skip empty emails
                            continue
                        # Basic email validation
                        if '@' not in email or '.' not in email:
                            print(f"Invalid email format: {email}")
                            continue
                        # Format for Monday.com API - send as a JSON object with email and text
                        column_values[col_id] = {
                            "email": email,
                            "text": email
                        }
                    elif col_type == "long_text":
                        column_values[col_id] = str(val)
                    else:
                        column_values[col_id] = str(val)
                except Exception as e:
                    print(f"Error processing column {col} value '{val}': {e}")
                    continue

            # Compose mutation
            query = """
            mutation ($boardId: ID!, $itemName: String!, $columnValues: JSON!) {
              create_item (board_id: $boardId, item_name: $itemName, column_values: $columnValues) {
                id
              }
            }
            """
            variables = {
                "boardId": board_id,
                "itemName": item_name,
                "columnValues": json.dumps(column_values)
            }

            response = requests.post(url, headers=headers, json={"query": query, "variables": variables})
            result = response.json()

            if 'errors' in result:
                print(f"Upload Error: {result}")
            else:
                if "create_item" in result["data"]:
                    print("\nSuccess! Item created with ID:", result["data"]["create_item"]["id"])
                elif "change_multiple_column_values" in result["data"]:
                    print("\nSuccess! Item updated with ID:", result["data"]["change_multiple_column_values"]["id"])
                else:
                    print("\nSuccess! Item processed.")
                success_count += 1

        self.status_label.config(text=f"Uploaded {success_count} items.")

    def open_table_editor(self):
        board_id = self.board_var.get()
        api_key = self.api_key_entry.get().strip()

        if not api_key or not board_id:
            messagebox.showwarning("Missing Info", "Please enter your API key and select a board.")
            return

        # Ensure we have user data
        if not self.users:
            self.fetch_users()
            if not self.users:
                messagebox.showerror("Error", "Failed to load users. Please check your API key and try again.")
                return

        # Fetch columns with settings
        query = """
        query ($boardId: [ID!]) {
        boards(ids: $boardId) {
            columns {
            id
            title
            type
            settings_str
            }
        }
        }
        """
        variables = { "boardId": [str(board_id)] }

        headers = {
            "Authorization": api_key,
            "Content-Type": "application/json"
        }

        response = requests.post("https://api.monday.com/v2", headers=headers,
                                json={"query": query, "variables": variables})
        result = response.json()

        if "errors" in result:
            messagebox.showerror("API Error", str(result["errors"]))
            return

        columns = result["data"]["boards"][0]["columns"]
        if not columns:
            messagebox.showwarning("No Columns", "No columns found on this board.")
            return

        # Map titles to IDs and store column settings
        column_titles = [col["title"] for col in columns]
        column_ids = [col["id"] for col in columns]
        column_settings = {}
        
        print("\nProcessing column settings:")
        for col in columns:
            if col["type"] == "status":
                print(f"\nStatus column: {col['title']}")
                print(f"Settings string: {col['settings_str']}")
                try:
                    settings = json.loads(col["settings_str"])
                    # Extract labels from the settings
                    if "labels" in settings:
                        # Convert labels to the correct format
                        labels = {}
                        for label in settings["labels"]:
                            if isinstance(label, dict) and "id" in label and "name" in label:
                                labels[str(label["id"])] = label["name"]
                        settings["labels"] = labels
                    column_settings[col["id"]] = settings
                    print(f"Parsed settings: {json.dumps(settings, indent=2)}")
                except Exception as e:
                    print(f"Error parsing settings: {e}")
                    # Default status options if settings can't be loaded
                    column_settings[col["id"]] = {
                        "labels": {
                            "0": "Active",
                            "1": "Independent"
                        }
                    }
                    print(f"Using default settings: {json.dumps(column_settings[col['id']], indent=2)}")

        # Spreadsheet UI
        editor = tk.Toplevel(self.root)
        editor.title("Spreadsheet Editor")
        
        # Get the main window size
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        
        # Set the editor window size to match main window
        editor.geometry(f"{main_width}x{main_height}")
        
        # Create a frame that fills the window
        main_frame = tk.Frame(editor)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights for the main frame
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        sheet = tksheet.Sheet(main_frame)
        sheet.enable_bindings((
            "single_select", "column_select", "row_select",
            "arrowkeys", "right_click_popup_menu",
            "rc_select", "rc_insert_row", "rc_delete_row",
            "rc_insert_column", "rc_delete_column",
            "copy", "cut", "paste", "delete", "undo", "edit_cell"
        ))
        sheet.grid(row=0, column=0, sticky="nsew")
        sheet.headers(column_titles)
        sheet.set_sheet_data([["" for _ in column_titles] for _ in range(10)])

        def create_dropdown(event, r, c):
            if c >= len(columns):
                return
                
            col_type = columns[c]["type"]
            col_id = columns[c]["id"]
            
            if col_type == "people":
                # Create dropdown window
                dropdown = tk.Toplevel(editor)
                dropdown.overrideredirect(True)
                
                # Get cell position
                x = sheet.winfo_rootx() + sheet.winfo_width() * c / len(columns)
                y = sheet.winfo_rooty() + sheet.winfo_height() * (r + 1) / 10
                dropdown.geometry(f"+{int(x)}+{int(y)}")
                
                # Create listbox with scrollbar
                frame = tk.Frame(dropdown)
                frame.pack(fill=tk.BOTH, expand=True)
                
                scrollbar = tk.Scrollbar(frame)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                
                listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set)
                listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                
                scrollbar.config(command=listbox.yview)
                
                # Add users to listbox
                user_names = sorted(self.users.keys())
                for name in user_names:
                    listbox.insert(tk.END, name)
                
                # Handle selection
                def on_select(event):
                    selected = [listbox.get(i) for i in listbox.curselection()]
                    if selected:
                        sheet.set_cell_data(r, c, ", ".join(selected))
                    dropdown.destroy()
                
                listbox.bind('<<ListboxSelect>>', on_select)
                listbox.bind('<Return>', on_select)
                
                # Close dropdown when clicking outside
                def on_click_outside(event):
                    if event.widget != listbox:
                        dropdown.destroy()
                
                dropdown.bind('<FocusOut>', on_click_outside)
                listbox.focus_set()
                
                # Set initial size
                dropdown.update_idletasks()
                width = max(200, dropdown.winfo_reqwidth())
                height = min(300, dropdown.winfo_reqheight())
                dropdown.geometry(f"{width}x{height}")
                
            elif col_type == "status":
                print(f"\nCreating status dropdown for column {col_id}")
                # Create dropdown window
                dropdown = tk.Toplevel(editor)
                dropdown.overrideredirect(True)
                
                # Get cell position
                x = sheet.winfo_rootx() + sheet.winfo_width() * c / len(columns)
                y = sheet.winfo_rooty() + sheet.winfo_height() * (r + 1) / 10
                dropdown.geometry(f"+{int(x)}+{int(y)}")
                
                # Create listbox
                listbox = tk.Listbox(dropdown, selectmode=tk.SINGLE)
                listbox.pack(fill=tk.BOTH, expand=True)
                
                # Add status options to listbox
                settings = column_settings.get(col_id, {})
                labels = settings.get("labels", {
                    "0": "Active",
                    "1": "Independent"
                })
                
                print(f"Available status labels: {json.dumps(labels, indent=2)}")
                
                for index, label in labels.items():
                    listbox.insert(tk.END, label)
                
                # Handle selection
                def on_select(event):
                    selected = listbox.get(listbox.curselection())
                    if selected:
                        print(f"Selected status: {selected}")
                        sheet.set_cell_data(r, c, selected)
                    dropdown.destroy()
                
                listbox.bind('<<ListboxSelect>>', on_select)
                listbox.bind('<Return>', on_select)
                
                # Close dropdown when clicking outside
                def on_click_outside(event):
                    if event.widget != listbox:
                        dropdown.destroy()
                
                dropdown.bind('<FocusOut>', on_click_outside)
                listbox.focus_set()

        # Bind double-click to create dropdown
        sheet.bind('<Double-Button-1>', lambda e: create_dropdown(e, sheet.get_selected_row(), sheet.get_selected_column()))

        # Store item IDs for each row
        item_ids = []

        def load_items():
            try:
                # Fetch items from the board
                query = """
                query ($boardId: [ID!]) {
                    boards(ids: $boardId) {
                        items_page {
                            items {
                                id
                                name
                                column_values {
                                    id
                                    value
                                    text
                                }
                            }
                        }
                    }
                }
                """
                response = requests.post("https://api.monday.com/v2", headers=headers,
                                      json={"query": query, "variables": variables})
                result = response.json()

                if "errors" in result:
                    messagebox.showerror("API Error", str(result["errors"]))
                    return

                items = result["data"]["boards"][0]["items_page"]["items"]
                if not items:
                    messagebox.showinfo("No Items", "No items found on this board.")
                    return

                # Clear existing data
                item_ids.clear()
                sheet_data = []

                # Process each item
                for item in items:
                    row_data = [""] * len(column_titles)
                    item_ids.append(item["id"])
                    
                    # Set the name
                    name_idx = next((i for i, col in enumerate(column_titles) if col.lower() == 'name'), None)
                    if name_idx is not None:
                        row_data[name_idx] = item["name"]

                    # Set other column values
                    for col_value in item["column_values"]:
                        col_idx = next((i for i, col in enumerate(columns) if col["id"] == col_value["id"]), None)
                        if col_idx is not None:
                            # Parse the value based on column type
                            col_type = columns[col_idx]["type"]
                            if col_type == "text":
                                row_data[col_idx] = col_value["text"]
                            elif col_type == "number":
                                try:
                                    row_data[col_idx] = float(col_value["value"])
                                except (ValueError, TypeError):
                                    row_data[col_idx] = col_value["text"]
                            elif col_type == "people":
                                # For people columns, show the names
                                try:
                                    people_data = json.loads(col_value["value"])
                                    if people_data and "personsAndTeams" in people_data:
                                        names = []
                                        for person in people_data["personsAndTeams"]:
                                            if person["kind"] == "person":
                                                # Find name by ID
                                                for name, user_id in self.users.items():
                                                    if str(user_id) == str(person["id"]):
                                                        names.append(name)
                                                        break
                                        row_data[col_idx] = ", ".join(names)
                                except:
                                    row_data[col_idx] = col_value["text"]
                            elif col_type == "status":
                                # For status columns, show the label
                                try:
                                    status_data = json.loads(col_value["value"])
                                    print(f"\nStatus data for {columns[col_idx]['title']}: {json.dumps(status_data, indent=2)}")
                                    if status_data and "label" in status_data:
                                        row_data[col_idx] = status_data["label"]
                                    else:
                                        row_data[col_idx] = col_value["text"]
                                except Exception as e:
                                    print(f"Error processing status value: {e}")
                                    row_data[col_idx] = col_value["text"]
                            else:
                                row_data[col_idx] = col_value["text"]

                    sheet_data.append(row_data)

                # Update the sheet
                sheet.set_sheet_data(sheet_data)
                messagebox.showinfo("Success", f"Loaded {len(items)} items")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to load items: {e}")

        def save_sheet():
            try:
                data = sheet.get_sheet_data()
                df = pd.DataFrame(data, columns=column_titles)
                output_path = "manual_upload.xlsx"
                df.to_excel(output_path, index=False)
                messagebox.showinfo("Saved", f"Spreadsheet saved to {output_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

        def upload_sheet():
            nonlocal item_ids  # Add this line to access the item_ids list
            try:
                data = sheet.get_sheet_data()
                df = pd.DataFrame(data, columns=column_titles)
                df.dropna(how='all', inplace=True)

                if df.empty:
                    messagebox.showwarning("Empty Data", "No data to upload.")
                    return

                success_count = 0
                total_rows = len(df)

                # Process each row
                for row_idx, row in df.iterrows():
                    # Check if row is empty (all values are empty strings or NaN)
                    if row.isna().all() or (row == "").all():
                        print(f"\nStopping at row {row_idx + 1} - empty row encountered")
                        break

                    print(f"\nProcessing row {row_idx + 1} of {total_rows}")
                    
                    # Find the name column (case-insensitive)
                    name_col = next((col for col in df.columns if col.lower() == 'name'), None)
                    item_name = str(row[name_col]) if name_col and not pd.isna(row[name_col]) else "Untitled Item"

                    # Create column values mapping
                    column_values = {}
                    print("\nProcessing column values:")
                    for idx, col in enumerate(columns):
                        col_id = col["id"]
                        col_type = col["type"]
                        col_title = col["title"]
                        val = row[column_titles[idx]]
                        
                        print(f"\nColumn: {col_title} (Type: {col_type}, ID: {col_id})")
                        print(f"Raw value: {val}")
                        
                        if pd.isna(val) or val == "":
                            print("Skipping empty value")
                            continue

                        try:
                            if col_type == "name":
                                column_values[col_id] = str(val)
                            elif col_type == "people":
                                # For people type, we need to send an array of user IDs
                                people_ids = []
                                for name in str(val).split(','):
                                    user_id = self.get_user_id(name.strip())
                                    if user_id:
                                        people_ids.append({"id": int(user_id), "kind": "person"})
                                column_values[col_id] = json.dumps({"personsAndTeams": people_ids})
                            elif col_type == "status":
                                # For status type, find the index for the label
                                settings = column_settings.get(col_id, {})
                                labels = settings.get("labels", {
                                    "0": "Active",
                                    "1": "Independent"
                                })
                                
                                print(f"Status settings for {col_title}: {json.dumps(settings, indent=2)}")
                                print(f"Available labels: {json.dumps(labels, indent=2)}")
                                
                                # Find the index for the label
                                status_index = None
                                for index, label in labels.items():
                                    if label.lower() == str(val).lower():
                                        status_index = int(index)
                                        break
                                
                                if status_index is not None:
                                    print(f"Found matching index {status_index} for label '{val}'")
                                    column_values[col_id] = json.dumps({
                                        "label": str(val),
                                        "index": status_index
                                    })
                                else:
                                    # If no matching label found, try to use the value as an index
                                    try:
                                        index = int(val)
                                        if str(index) in labels:
                                            print(f"Using numeric index {index}")
                                            column_values[col_id] = json.dumps({
                                                "label": labels[str(index)],
                                                "index": index
                                            })
                                        else:
                                            print(f"Invalid status index: {index}")
                                            continue
                                    except ValueError:
                                        print(f"Invalid status value: {val}")
                                        continue
                            elif col_type == "date":
                                if isinstance(val, pd.Timestamp):
                                    val = val.strftime('%Y-%m-%d')
                                column_values[col_id] = {"date": str(val)}
                            elif col_type == "board_relation":
                                item_ids = [int(id.strip()) for id in str(val).split(",") if id.strip()]
                                column_values[col_id] = json.dumps({"item_ids": item_ids})
                            elif col_type == "text":
                                column_values[col_id] = str(val)
                            elif col_type == "phone":
                                # Format phone number (remove any non-digit characters)
                                phone = ''.join(filter(str.isdigit, str(val)))
                                column_values[col_id] = phone
                            elif col_type == "email":
                                # Email needs to be sent as a JSON object with email and text
                                email = str(val).strip()
                                if not email:  # Skip empty emails
                                    continue
                                # Basic email validation
                                if '@' not in email or '.' not in email:
                                    print(f"Invalid email format: {email}")
                                    continue
                                # Format for Monday.com API - send as a JSON object with email and text
                                column_values[col_id] = {
                                    "email": email,
                                    "text": email
                                }
                            elif col_type == "long_text":
                                column_values[col_id] = str(val)
                            elif col_type == "number":
                                try:
                                    num_val = float(val)
                                    column_values[col_id] = num_val
                                except (ValueError, TypeError):
                                    print(f"Failed to convert to number: {val}")
                                    continue
                            else:
                                # Default to text for unknown types
                                column_values[col_id] = str(val)
                                
                            print(f"Formatted value: {column_values[col_id]}")
                        except Exception as e:
                            print(f"Error processing column {col_title}: {str(e)}")
                            continue

                    print("\nFinal column_values:", json.dumps(column_values, indent=2))

                    # Check if this is an existing item or new item
                    item_id = item_ids[row_idx] if row_idx < len(item_ids) else None
                    
                    if item_id:
                        # Update existing item (including name)
                        column_values["name"] = item_name  # Set the name in the column values
                        query = """
                        mutation ($itemId: ID!, $boardId: ID!, $columnValues: JSON!) {
                          change_multiple_column_values (item_id: $itemId, board_id: $boardId, column_values: $columnValues) {
                            id
                          }
                        }
                        """
                        variables = {
                            "itemId": item_id,
                            "boardId": board_id,
                            "columnValues": json.dumps(column_values)
                        }
                    else:
                        # Create new item
                        query = """
                        mutation ($boardId: ID!, $itemName: String!, $columnValues: JSON!) {
                          create_item (board_id: $boardId, item_name: $itemName, column_values: $columnValues) {
                            id
                          }
                        }
                        """
                        variables = {
                            "boardId": board_id,
                            "itemName": item_name,
                            "columnValues": json.dumps(column_values)
                        }

                    headers = {
                        "Authorization": api_key,
                        "Content-Type": "application/json"
                    }

                    print("\nSending request to Monday.com API...")
                    response = requests.post("https://api.monday.com/v2", headers=headers, 
                                          json={"query": query, "variables": variables})
                    result = response.json()
                    
                    if 'errors' in result:
                        print("\nAPI Error:", json.dumps(result["errors"], indent=2))
                        messagebox.showerror("Upload Error", str(result["errors"]))
                    else:
                        if "create_item" in result["data"]:
                            print("\nSuccess! Item created with ID:", result["data"]["create_item"]["id"])
                        elif "change_multiple_column_values" in result["data"]:
                            print("\nSuccess! Item updated with ID:", result["data"]["change_multiple_column_values"]["id"])
                        else:
                            print("\nSuccess! Item processed.")
                        success_count += 1

                messagebox.showinfo("Success", f"Successfully uploaded {success_count} of {total_rows} items")
                    
            except Exception as e:
                print("\nUnexpected error:", str(e))
                messagebox.showerror("Upload Error", f"Failed to upload data: {e}")

        # Add buttons
        button_frame = tk.Frame(main_frame)
        button_frame.grid(row=1, column=0, sticky="ew", pady=5)
        
        tk.Button(button_frame, text="Load Items", command=load_items).pack(side="left", padx=5)
        tk.Button(button_frame, text="Save to Excel", command=save_sheet).pack(side="left", padx=5)
        tk.Button(button_frame, text="Upload to Monday.com", command=upload_sheet).pack(side="right", padx=5)

# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = MondayUploaderApp(root)
    root.mainloop()
