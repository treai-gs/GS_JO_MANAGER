import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import asana
from datetime import date, datetime, timezone
import argparse
import sys, os
import json
import xml.dom.minidom as xr
import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta
import requests
import time
import threading



os.environ['ASANA_CLIENT_ID'] = "1203432864266713"
os.environ['ASANA_CLIENT_SECRET'] = "ff2f2249e9813c271ec877d4345ace4f"

# Important Asana GIDs
team_gid = "735937443511725" # Geospatial Solutions Italy
proc_gid = "1167558511333185" # Processing
template_gid_list = ["1202547950897955", "1203588081135756"] # templates
projects_overview_gid = "1202122023599837" # GS master project


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # configure the root window
        self.title('GS Job Order Manager')
        self.geometry('1050x800')

# Frame 1: Login to Asana
        self.frame1 = tk.LabelFrame(self, text="Asana Sign-In", width=450, height=300)
        self.frame1.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")
        self.frame1.grid_propagate(False)

        # Button: Get Login Code
        self.get_code_button = ttk.Button(self.frame1, text='Get Asana Login Code', command=self.asana_login)
        self.get_code_button.grid(row = 1, column = 0, columnspan=2, padx=10, pady=10,sticky="ew")

        # Label: Login Code
        self.code_label = ttk.Label(self.frame1, text="Enter Login Code:")
        self.code_label.grid(row = 2, column = 0, padx=10, pady=10)

        # Entry: Login Code
        self.code_entry = ttk.Entry(self.frame1, width=45)
        self.code_entry.grid(row = 2, column = 1, padx=10, pady=10, sticky="ew")

        # Button: Sign in
        self.login_button = ttk.Button(self.frame1, text='Login', command=lambda: (self.asana_auth(), self.asana_get_gids()))
        self.login_button.grid(row = 3, column = 0, columnspan=2, padx=10, pady=10,sticky="ew")

        # Label: Login Code
        self.status_label = ttk.Label(self.frame1, text="Not logged in.")
        self.status_label.grid(row = 4, column = 0, columnspan=2, padx=10, pady=10)

    # Frame 2: External JO Information
        self.frame2 = tk.LabelFrame(self, text="External JO Information", width=450, height=450)
        self.frame2.grid(row=1, column=0, padx=20, sticky="nsew")
        self.frame2.grid_propagate(False)

        # Button: Select XML
        self.upload_xml_button = ttk.Button(self.frame2, text='Select XML', command=self.load_xml_path)
        self.upload_xml_button.grid(row = 0, column = 0, padx=10, pady=10,sticky="ew")

        # Entry: XML path
        self.xml_path = ttk.Entry(self.frame2, width=45)
        self.xml_path.insert(0, "XML Path")
        self.xml_path.grid(row = 0, column = 1, sticky="ew")

        # Scrollbar: XML path
        self.scroll_xml_path = ttk.Scrollbar(self.frame2, orient='horizontal', command=self.xml_path.xview, )
        self.xml_path.config(xscrollcommand=self.scroll_xml_path.set)
        self.scroll_xml_path.grid(row = 1, column = 1, sticky='ew')

        # Button: Load XML 
        self.load_xml_button = ttk.Button(self.frame2, text='Load XML Data', command=self.parse_xml)
        self.load_xml_button.grid(row = 2, column = 0, columnspan=2, padx=10, pady=10,sticky="ew")

        # Label: JO
        self.JO_label = ttk.Label(self.frame2, text="Job Order:")
        self.JO_label.grid(row = 3, column = 0, padx=10, pady=10, sticky="w")

        # Label: JO value
        self.JO_val = ttk.Label(self.frame2, width=45)
        self.JO_val.grid(row = 3, column = 1, sticky="w")

        # Label: Start Date
        self.start_label = ttk.Label(self.frame2, text="Expected Init:")
        self.start_label.grid(row = 4, column = 0, padx=10, pady=10, sticky="w")

        # Label: Start Date value
        self.start_val = ttk.Label(self.frame2, width=45)
        self.start_val.grid(row = 4, column = 1, sticky="w")

        # Label: End Date
        self.proc_end_label = ttk.Label(self.frame2, text="Expected Processing End:")
        self.proc_end_label.grid(row = 5, column = 0, padx=10, pady=10,sticky="w")

        # Label: End Date value
        self.proc_end_val = ttk.Label(self.frame2, width=45)
        self.proc_end_val.grid(row = 5, column = 1, sticky="w")

        # Label: Delivery Date
        self.delivery_end_label = ttk.Label(self.frame2, text="Delivery date:")
        self.delivery_end_label.grid(row = 6, column = 0, padx=10, pady=10, sticky="w")

        # Label: End Date value
        self.delivery_end_val = ttk.Label(self.frame2, width=45)
        self.delivery_end_val.grid(row = 6, column = 1, sticky="w")

    # Frame 3: Configure project
        self.frame3 = tk.LabelFrame(self, text="Configure Project", width=500, height=750)
        self.frame3.grid(row=0, column=2, rowspan=2, padx=20, pady=10, sticky="nsew")
        self.frame3.grid_propagate(False)

        # Label: Project Title
        self.prj_title_label = ttk.Label(self.frame3, text="Project Title:")
        self.prj_title_label.grid(row = 0, column = 0, padx=10, pady=10, sticky="w")

        # Entry: Project Title
        self.prj_title_entry = ttk.Entry(self.frame3, width=45)
        self.prj_title_entry.grid(row = 0, column = 1, sticky="w")

        # Label: Project Label
        self.prj_label_label = ttk.Label(self.frame3, text="Project Label:")
        self.prj_label_label.grid(row = 1, column = 0, padx=10, pady=10, sticky="w")

        # Entry: Project Label
        self.prj_label_entry = ttk.Entry(self.frame3, width=45)
        self.prj_label_entry.grid(row = 1, column = 1, sticky="w")

        # Label: Start Date
        self.start_label = ttk.Label(self.frame3, text="Start Date (YYYY-mm-dd):")
        self.start_label.grid(row = 2, column = 0, padx=10, pady=10, sticky="w")

        # Entry: Start Date
        self.start_entry = ttk.Entry(self.frame3, width=20)
        self.start_entry.grid(row = 2, column = 1, sticky="w")

        # Label: Delivery Start Date
        self.d_start_label = ttk.Label(self.frame3, text="First Update (YYYY-mm-dd):")
        self.d_start_label.grid(row = 3, column = 0, padx=10, pady=10, sticky="w")

        # Entry: Delivery Start Date
        self.d_start_entry = ttk.Entry(self.frame3, width=20)
        self.d_start_entry.grid(row = 3, column = 1, sticky="w")


        # # Label: End Date
        # self.end_label = ttk.Label(self.frame3, text="End Date (YYYY-mm-dd):")
        # self.end_label.grid(row = 2, column = 0, padx=10, pady=10, sticky="w")

        # # Entry: End Date
        # self.end_entry = ttk.Entry(self.frame3, width=45)
        # self.end_entry.grid(row = 2, column = 1, sticky="w")

        # Label: Number of updates
        self.n_updates_label = ttk.Label(self.frame3, text="Number of Updates:")
        self.n_updates_label.grid(row = 4, column = 0, padx=10, pady=10, sticky="w")

        # Entry: Number of updates
        self.n_updates_entry = ttk.Entry(self.frame3, width=10)
        self.n_updates_entry.grid(row = 4, column = 1, sticky="w")

        # Label: Update Time Interval
        self.update_int_label = ttk.Label(self.frame3, text="Update Time Interval (days):")
        self.update_int_label.grid(row = 5, column = 0, padx=10, pady=10, sticky="w")

        # Entry: Update Time Interval
        self.update_int = ttk.Entry(self.frame3, width=10)
        self.update_int.insert(0, "1")
        self.update_int.grid(row = 5, column = 1, sticky="w")

        # Label: Select Template
        self.template_label = ttk.Label(self.frame3, text="Template")
        self.template_label.grid(row = 6, column = 0, padx=10, pady=10, sticky="w")

        # Combobox: Select Template
        self.template_combo_text = tk.StringVar()
        self.template_combo = ttk.Combobox(self.frame3, width=45, textvariable=self.template_combo_text, state="readonly")
        self.template_combo.grid(row = 6, column = 1, sticky="w")

        # Label: Select Deliverable
        self.deliverable_label = ttk.Label(self.frame3, text="Deliverable")
        self.deliverable_label.grid(row = 7, column = 0, padx=10, pady=10, sticky="w")

        # Menu: Select Deliverable
        
        self.deliverable_squeesar1d = tk.BooleanVar()
        self.deliverable_squeesar2d = tk.BooleanVar()
        self.deliverable_changedet = tk.BooleanVar()
        self.deliverable_rmt = tk.BooleanVar()
        self.deliverable_rmt3d = tk.BooleanVar()
        self.deliverable_fullrep = tk.BooleanVar()
        self.deliverable_smallrep = tk.BooleanVar()

        self.deliverable_menu = tk.Menu(self.frame3, tearoff=0)
        self.deliverable_button = ttk.Menubutton(self.frame3, text="Select Deliverables", menu=self.deliverable_menu)
        self.deliverable_menu.add_checkbutton(label="SqueeSAR 1D", onvalue=1, offvalue=0, variable=self.deliverable_squeesar1d)
        self.deliverable_menu.add_checkbutton(label="SqueeSAR 2D", onvalue=1, offvalue=0, variable=self.deliverable_squeesar2d)
        self.deliverable_menu.add_checkbutton(label="Change Detection", onvalue=1, offvalue=0, variable=self.deliverable_changedet)
        self.deliverable_menu.add_checkbutton(label="RMT", onvalue=1, offvalue=0, variable=self.deliverable_rmt)
        self.deliverable_menu.add_checkbutton(label="RMT 3D", onvalue=1, offvalue=0, variable=self.deliverable_rmt3d)
        self.deliverable_menu.add_checkbutton(label="Full Report", onvalue=1, offvalue=0, variable=self.deliverable_fullrep)
        self.deliverable_menu.add_checkbutton(label="Short Report", onvalue=1, offvalue=0, variable=self.deliverable_smallrep)

        # self.deliverable_menu.add_cascade(label='View', menu=self.deliverable_menu)
        # self.deliverable_combo = ttk.Combobox(self.frame3, width=45, textvariable=self.template_combo_text, state="readonly")
        self.deliverable_button.grid(row = 7, column = 1, sticky="w")        

        # Label: Bulletin checkbutton
        self.bulletin_check_label = ttk.Label(self.frame3, text="Bulletin")
        self.bulletin_check_label.grid(row = 8, column = 0, padx=10, pady=10, sticky="w")     

        # Checkbutton: Bulletin checkbutton
        self.bulletin_check_button_value = tk.IntVar()
        self.bulletin_check_button = ttk.Checkbutton(self.frame3, variable=self.bulletin_check_button_value, text="Include", state='disabled')
        self.bulletin_check_button.grid(row = 8, column = 1, sticky="w")
        self.template_combo.bind("<<ComboboxSelected>>", self.toggle_checkbutton)        

        # Label: Select Technical Responsible
        self.tr_label = ttk.Label(self.frame3, text="Technical Responsible:")
        self.tr_label.grid(row = 9, column = 0, padx=10, pady=10, sticky="w")

        # Combobox: Select Technical Responsible
        self.tr_combo_text = tk.StringVar()
        self.tr_combo = ttk.Combobox(self.frame3, width=45, textvariable=self.tr_combo_text, state="readonly")
        self.tr_combo.grid(row = 9, column = 1, sticky="w")

        # Label: Select Processing Responsible
        self.pr_label = ttk.Label(self.frame3, text="Processing Operator:")
        self.pr_label.grid(row = 10, column = 0, padx=10, pady=10, sticky="w")

        # Combobox: Select Processing Responsible
        self.pr_combo_text = tk.StringVar()
        self.pr_combo = ttk.Combobox(self.frame3, width=45, textvariable=self.pr_combo_text, state="readonly")
        self.pr_combo.grid(row = 10, column = 1, sticky="w")

        # Button: Submit JO to Asana
        self.submit_button = ttk.Button(self.frame3, text='Submit JO to Asana', command=self.start_submit_thread)
        self.submit_button.grid(row = 11, column = 0, columnspan=2, padx=10, pady=10,sticky="ew")

        # Progress bar: Asana
        self.asana_pb = ttk.Progressbar(self.frame3, orient="horizontal", mode="indeterminate")
        self.asana_pb.grid(row = 12, column = 0, columnspan=2, padx=10, pady=10,sticky="ew")

        # Label: Asana progress bar
        self.label_pb = ttk.Label(self.frame3)
        self.label_pb.grid(row = 13, column = 0, columnspan=2, padx=10)

# Functions
    # Load XML file from NAV
    def load_xml_path(self):
        fp = tk.filedialog.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=(('XML Files', '*.xml'),))
        self.xml_path.delete(0, 'end')
        self.xml_path.insert(0, fp)
        return
    
    # Generate access token for user
    def asana_login(self):
        if 'ASANA_CLIENT_ID' in os.environ:
            # create a client with the OAuth credentials:
            self.client = asana.Client.oauth(
                client_id=os.environ['ASANA_CLIENT_ID'],
                client_secret=os.environ['ASANA_CLIENT_SECRET'],
                # this special redirect URI will prompt the user to copy/paste the code.
                # useful for command line scripts and other non-web apps
                redirect_uri='urn:ietf:wg:oauth:2.0:oob'
            )
            print("authorized=", self.client.session.authorized)
            # get an authorization URL:
            (url, state) = self.client.session.authorization_url()
            try:
                # in a web app you'd redirect the user to this URL when they take action to
                # login with Asana or connect their account to Asana
                import webbrowser
                webbrowser.open(url)
            except Exception as e:
                print("Open the following URL in a browser to authorize:")
                print(url)
    
    # Login with the access token
    def asana_auth(self):
        code = self.code_entry.get()
        # exchange the code for a bearer token
        token = self.client.session.fetch_token(code=code)

        # normally you'd persist this token somewhere
        os.environ['ASANA_TOKEN'] = json.dumps(token) # (see below)
        self.me = self.client.users.me()

        self.status_label["text"] = "Logged in as %s" % (self.me["name"])



    def asana_get_gids(self):
        self.users_gs = pd.DataFrame(self.client.users.get_users_for_team(team_gid))
        self.users_proc = pd.DataFrame(self.client.users.get_users_for_team(proc_gid))
        self.tr_combo["values"] = self.users_gs["name"].to_list()
        self.tr_combo_text.set(self.me["name"])
        self.pr_combo["values"] = self.users_proc["name"].to_list()
        self.template_combo["values"] = ["Supervised", "Unsupervised"]
        self.template_combo_text.set("Supervised")

    def parse_xml(self):
        # dates in NAV changed format at some point: try different format to read dates from NAV xml
        def get_date(date_str):
            date_formats = ["%y-%m-%d", "%d/%m/%y", "%Y-%m-%d","%Y/%m/%d","%m/%d/%y"]
            for date_format in date_formats:
                try:
                    return datetime.strptime( date_str, date_format )
                except:
                    pass
        # Parse xml
        xml = xr.parse(self.xml_path.get())
        # Get some attributes: Job Order number and delivery date (to the client), expected processing init and end
        JOnumber = xml.getElementsByTagName("JobOrderCode")[0].firstChild.nodeValue
        description = xml.getElementsByTagName("Description")[0].firstChild.nodeValue
        delivery_code = xml.getElementsByTagName("Delivery")[0].getElementsByTagName("Code")[0].firstChild.nodeValue
        delivery_date = xml.getElementsByTagName("Delivery")[0].getElementsByTagName("Date")[0].firstChild.nodeValue
        # All dates are strings, change to datetime format
        delivery_date = datetime.strptime(delivery_date[:10], "%Y-%m-%d")
        expected_init = get_date(xml.getElementsByTagName("DeliveryExpectedInit")[0].firstChild.nodeValue)
        expected_end = get_date(xml.getElementsByTagName("DeliveryExpectedProcessingEnd")[0].firstChild.nodeValue)

        self.title("GS Job Order Manager: [JO%s] %s" % (JOnumber, description))
        self.JO_val["text"] = "[JO%s] %s" % (JOnumber, description)
        self.start_val["text"] = expected_init.strftime("%Y-%m-%d")
        self.proc_end_val["text"] = expected_end.strftime("%Y-%d-%m")
        self.delivery_end_val["text"] = delivery_date.strftime("%Y-%m-%d")

        self.prj_title_entry.delete(0, 'end')
        self.prj_title_entry.insert(0, "[JO%s] %s" % (JOnumber, description))
        self.prj_label_entry.delete(0, 'end')
        self.prj_label_entry.insert(0, "[JO%s] %s" % (JOnumber, description))
        self.start_entry.delete(0, 'end')
        self.start_entry.insert(0, self.start_val["text"])
        self.d_start_entry.delete(0, 'end')
        self.d_start_entry.insert(0, self.start_val["text"])

    def toggle_checkbutton(self, event):
        print("executed")
        if self.template_combo_text.get() == "Supervised":
            self.bulletin_check_button_value.set(0)
            self.bulletin_check_button.configure(state='disabled')
        else:
            self.bulletin_check_button.configure(state='normal')

    def submit_to_asana(self):
        # Duplicate the template project
        self.label_pb["text"] = "Duplicating template project..."
        self.frame3.update_idletasks()
        self.asana_pb.start([25])
        if self.template_combo_text.get() == "Supervised":
            template_gid = template_gid_list[0]
        else:
            template_gid = template_gid_list[1]
        dup_project = self.client.projects.duplicate_project(template_gid, {
                "include": [
                    "task_subtasks",
                    "task_attachments",
                    "task_dates",
                    "task_dependencies",
                    "task_followers",
                    "task_projects"
                ],
                "name": self.prj_title_entry.get(),
                "schedule_dates": {
                    "start_on": self.start_entry.get(),
                    "should_skip_weekends": False
                },
            })

        project_gid = dup_project["new_project"]["gid"]
        print(project_gid)

        time.sleep(150) # Asana is slow to create the project, so unfortunately this must be programmed in. Consider in future checking repeatedly until a certain # of tasks are there

        self.label_pb["text"] = "Assigning tasks to users..."
        self.frame3.update_idletasks()

        # Assign tasks to proper team members
        assignee_gs = self.users_gs[self.users_gs["name"] == self.tr_combo_text.get()]["gid"].values[0] # Assignee for GS
        assignee_proc = self.users_proc[self.users_proc["name"] == self.pr_combo_text.get()]["gid"].values[0] # Placeholder assignee for proc
        
        # The tasks that need to be assigned to Processing. In the future, consider using the Team field
        if self.template_combo_text.get() == "Supervised":
            proc_task_names = ["Data Processing", "Processing completed", "Data Finalization", "Deliverables Generation", "Update 1- Data Processing", "Update 1- Processing completed", "Update 1-  Data Finalization", "Update 1- Deliverables Generation"]
        else:
            proc_task_names = ["Data Processing", "Processing completed", "Data Finalization", "Deliverables Generation", "Update 1- Data Processing", "Update 1- Processing completed", "Update 1- Imagery completed"]

        if self.bulletin_check_button_value.get() == 1:
            proc_task_names.append("Update 1- Data Processing - Bulletin")
            proc_task_names.append("Data Processing - Bulletin") 
        

        project_tasks = pd.DataFrame(self.client.tasks.get_tasks_for_project(project_gid, {'opt_fields': ['gid', 'name', 'start_on', 'due_on']}))
        # if self.bulletin_check_button_value.get() == 0:
        #     project_tasks.drop(project_tasks[project_tasks["name"].str.contains("Bulletin")].index, inplace=True)
        print(project_tasks)
        # project_tasks
        for task_gid, task_name, task_start, task_end in zip(project_tasks["gid"], project_tasks["name"], project_tasks["start_on"], project_tasks["due_on"]):
            if self.bulletin_check_button_value.get() == 0 and "Bulletin" in task_name:
                self.client.tasks.delete_task(task_gid)
                print("deleted", task_gid)
            # ELSE EVERYTHING BELOW
            else:
                print(task_gid)
                new_name = self.prj_label_entry.get() + " " + task_name
                proc_params = {'assignee': assignee_proc, 'name': new_name}
                gs_params = {'assignee': assignee_gs, 'name': new_name}
                if (int(self.update_int.get()) > 1) and ("Update" in task_name):
                    try:
                        new_start = datetime.strptime(task_start, "%Y-%m-%d") + relativedelta(days=int(self.update_int.get())-1)
                        new_start = new_start.strftime("%Y-%m-%d")
                    except TypeError:
                        new_start = ""
                    try:
                        new_end = datetime.strptime(task_end, "%Y-%m-%d") + relativedelta(days=int(self.update_int.get())-1)
                        new_end = new_end.strftime("%Y-%m-%d")
                    except TypeError:
                        new_end = ""
                    proc_params = {'assignee': assignee_proc, 'name': new_name, "start_on": new_start, "due_on": new_end}
                    gs_params = {'assignee': assignee_gs, 'name': new_name, "start_on": new_start, "due_on": new_end}
                if task_name in proc_task_names:
                    self.client.tasks.update_task(task_gid, proc_params)
                else:
                    self.client.tasks.update_task(task_gid, gs_params)
                    # self.client.tasks.add_project_for_task(task_gid, {'project': projects_overview_gid})
        
        # Create number of sections based on n_updates
        n_updates = int(self.n_updates_entry.get())
        section_data = list(self.client.sections.get_sections_for_project(project_gid=project_gid)) # Get section data for the project
        df_section = pd.DataFrame(section_data) # Convert to df
        update_gid = df_section[df_section["name"] == "Update 1"]["gid"].values[0] # GID of Update 1 which will be used as a template Section for the rest of the updates
        update_tasks = list(self.client.tasks.get_tasks_for_section(section_gid=update_gid))
        df_update_tasks_all = pd.DataFrame(update_tasks)
        if self.bulletin_check_button_value.get() == 0:
            df_update_tasks = df_update_tasks_all.drop(df_update_tasks_all[df_update_tasks_all["name"].str.contains("Bulletin")].index)
        else:
            df_update_tasks = df_update_tasks_all
        df_update_tasks["name"] = df_update_tasks["name"].str.replace("Update 1- ", "")

        # Iterate through each section in reverse order
        if int(self.n_updates_entry.get()) > 1:
            for n in reversed(range(1, n_updates+1)):
                self.label_pb["text"] = "Creating update %s..." % (str(n))
                self.frame3.update_idletasks()
                print(n)
                is_first_task=True
                new_section = self.client.sections.create_section_for_project(project_gid, {'name': 'Update %s' % (str(n)), 'insert_after': update_gid})
                for task_gid, task_name in zip(df_update_tasks["gid"],df_update_tasks["name"]):
                    dup_task = self.client.tasks.duplicate_task(task_gid, {'name': 'Update %s- %s' % (str(n), task_name), "include": ["dates", "assignee", "projects", "subtasks"]})
                    new_start = datetime.strptime(self.d_start_entry.get(), "%Y-%m-%d")


                    try:
                        old_start = self.client.tasks.get_task(dup_task["new_task"]["gid"])["start_on"] # default start date of the task 

                        if is_first_task is True: # if it is the first task, get the offset between the first task and the new start date for first task
                            old_start_t0 = datetime.strptime(old_start,"%Y-%m-%d")
                            offset_time = relativedelta(new_start, old_start_t0)
                            is_first_task = False

                        
                         
                        update_int_offset = relativedelta(days=(n-1)*int(self.update_int.get()))
                        # start_on = datetime.strptime(self.client.tasks.get_task(dup_task["new_task"]["gid"])["start_on"], "%Y-%m-%d") + relativedelta(months=(n-1)*int(self.update_int.get()))
                        start_on = datetime.strptime(old_start, "%Y-%m-%d") + offset_time + update_int_offset
                        start_on = start_on.strftime("%Y-%m-%d")
                    
                    except TypeError:
                        start_on = ""
                    try:
                        old_end = self.client.tasks.get_task(dup_task["new_task"]["gid"])["due_on"] # default end date of the task
                        if is_first_task is True: # if it is the first task, get the offset between the first task and the new start date for first task
                            old_end_t0 = datetime.strptime(old_start, "%Y-%m-%d")
                            offset_time = relativedelta(new_start, old_end_t0)
                            is_first_task = False                        
                        # due_on = datetime.strptime(self.client.tasks.get_task(dup_task["new_task"]["gid"])["due_on"], "%Y-%m-%d") + relativedelta(months=(n-1)*int(self.update_int.get()))
                        update_int_offset = relativedelta(days=(n-1)*int(self.update_int.get()))
                        due_on = datetime.strptime(old_end, "%Y-%m-%d") + offset_time + update_int_offset
                        due_on = due_on.strftime("%Y-%m-%d")
                    except TypeError:
                        due_on = ""
                    self.client.tasks.update_task(dup_task["new_task"]["gid"], {'start_on': start_on, 'due_on': due_on})
                    self.client.tasks.add_project_for_task(dup_task["new_task"]["gid"], {'project': project_gid, 'section': new_section["gid"]},)
        # elif int(self.n_updates_entry.get()) == 0: # need to delete update 1 no matter what
        self.label_pb["text"] = "Removing update 1..."
        self.frame3.update_idletasks()
        for task_gid in df_update_tasks_all["gid"]:
            self.client.tasks.delete_task(task_gid)
        self.client.sections.delete_section(update_gid)

        self.asana_pb.stop()
        self.label_pb["text"] = "Asana project created."

    def start_submit_thread(self):
        global submit_thread
        submit_thread = threading.Thread(target=self.submit_to_asana)
        submit_thread.daemon = True
        submit_thread.start()
        app.after(20, self.check_submit_thread)
    
    def check_submit_thread(self):
        if submit_thread.is_alive():
            app.after(20, self.check_submit_thread)


if __name__ == "__main__":
  app = App()
  app.protocol("WM_DELETE_WINDOW", sys.exit)
  app.mainloop()
