"""Simple Tkinter desktop interface for the class planning assignment."""

import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from simple_engine import check_excel_plan, export_plan, generate_plan

PROJECT_FOLDER = Path(__file__).resolve().parent
PROGRAM_NAME = "Computer Science (B.S.) - Software Systems Track"
NAVY = "#233754"
BLUE = "#245CCB"
PALE = "#F3F6FB"
WHITE = "#FFFFFF"
TEXT = "#213047"
MUTED = "#63718A"


class SmartPlanningApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Smart Class Planning Tool")
        self.geometry("1050x690")
        self.minsize(880, 580)
        self.configure(bg=PALE)

        self.pdf_file = tk.StringVar(value="No DegreeWorks PDF selected")
        self.start_year = tk.StringVar(value="2026")
        self.course_limit = tk.StringVar(value="4")
        self.program = tk.StringVar(value=PROGRAM_NAME)
        self.summary = tk.StringVar(value="Choose a DegreeWorks PDF and click Generate Class Plan.")
        self.status = tk.StringVar(value="Ready")
        self.excel_path = None
        self.generated_result = None

        self._setup_style()
        self._build_interface()

    def _setup_style(self):
        # Sets consistent fonts and colors for the desktop controls.
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Course.Treeview", rowheight=29, font=("Segoe UI", 10),
                        background=WHITE, fieldbackground=WHITE, foreground=TEXT)
        style.configure("Course.Treeview.Heading", font=("Segoe UI", 10, "bold"),
                        background="#E9EFF8", foreground=TEXT)
        style.configure("TNotebook", background=WHITE, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(15, 10), font=("Segoe UI", 10))

    def _text(self, parent, label, size=10, bold=False, color=TEXT, **kwargs):
        # Creates labels with consistent desktop typography.
        return tk.Label(parent, text=label, bg=WHITE, fg=color,
                        font=("Segoe UI", size, "bold" if bold else "normal"),
                        anchor="w", **kwargs)

    def _button(self, parent, label, command, primary=False):
        # Creates a standard button using the application's color palette.
        return tk.Button(parent, text=label, command=command, cursor="hand2",
                         font=("Segoe UI", 10, "bold"), padx=14, pady=9,
                         relief="flat", bd=0,
                         bg=BLUE if primary else "#E7EEFA",
                         fg=WHITE if primary else "#1D498B",
                         activebackground="#18489F" if primary else "#D7E5FA",
                         activeforeground=WHITE if primary else "#1D498B")

    def _build_interface(self):
        # Places the program setup beside the results and Excel-checking tabs.
        header = tk.Frame(self, bg=NAVY, height=94)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="SMART CLASS PLANNING TOOL", bg=NAVY, fg=WHITE,
                 font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=28, pady=(14, 0))
        tk.Label(header, text="Generate a schedule or check an existing plan",
                 bg=NAVY, fg="#CBD8EC", font=("Segoe UI", 10)).pack(anchor="w", padx=29)

        footer = tk.Label(self, textvariable=self.status, bg="#E6ECF6", fg=MUTED,
                          font=("Segoe UI", 9), anchor="w", padx=20, pady=8)
        footer.pack(side="bottom", fill="x")

        body = tk.Frame(self, bg=PALE)
        body.pack(expand=True, fill="both", padx=18, pady=16)
        body.grid_columnconfigure(0, weight=0, minsize=300)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        left = tk.Frame(body, bg=WHITE, width=320, highlightthickness=1, highlightbackground="#DDE5F1")
        left.pack_propagate(False)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        right = tk.Frame(body, bg=WHITE, highlightthickness=1, highlightbackground="#DDE5F1")
        right.grid(row=0, column=1, sticky="nsew")
        self._build_setup(left)
        self._build_results(right)

    def _build_setup(self, parent):
        # Keeps the Generate button visible while maintaining compact setup controls.
        self._button(parent, "Generate Class Plan", self._generate, primary=True).pack(
            side="bottom", fill="x", padx=22, pady=(8, 22))
        # Allows setup fields to scroll without moving the Generate button.
        holder = tk.Frame(parent, bg=WHITE)
        holder.pack(side="top", fill="both", expand=True, padx=22, pady=(16, 4))
        canvas = tk.Canvas(holder, bg=WHITE, highlightthickness=0)
        scrollbar = ttk.Scrollbar(holder, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        panel = tk.Frame(canvas, bg=WHITE)
        window_id = canvas.create_window((0, 0), window=panel, anchor="nw")
        panel.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))
        self._text(panel, "Plan setup", 16, bold=True).pack(anchor="w", pady=(0, 14))

        self._text(panel, "DEGREE PROGRAM", 9, bold=True, color=MUTED).pack(anchor="w")
        self._button(panel, "Search degree programs...", self._choose_program).pack(
            fill="x", pady=(8, 6))
        tk.Label(panel, textvariable=self.program, bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9), wraplength=260, justify="left",
                 anchor="w").pack(fill="x", pady=(0, 12))

        self._text(panel, "DEGREEWORKS PDF", 9, bold=True, color=MUTED).pack(anchor="w")
        self._button(panel, "Choose PDF...", self._choose_pdf).pack(fill="x", pady=(8, 6))
        tk.Label(panel, textvariable=self.pdf_file, bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 9), wraplength=260, justify="left",
                 anchor="w").pack(fill="x", pady=(0, 7))
        tk.Button(panel, text="Use included sample PDF", command=self._sample_pdf,
                  bg=WHITE, fg=BLUE, font=("Segoe UI", 9, "underline"),
                  relief="flat", bd=0, cursor="hand2", anchor="w").pack(fill="x")

        self._text(panel, "PLAN SETTINGS", 9, bold=True, color=MUTED).pack(
            anchor="w", pady=(12, 7))
        settings = tk.Frame(panel, bg=WHITE)
        settings.pack(fill="x")
        settings.grid_columnconfigure(0, weight=1)
        settings.grid_columnconfigure(1, weight=1)
        self._text(settings, "Start year", 9, color=MUTED).grid(row=0, column=0, sticky="w")
        self._text(settings, "Max courses / term", 9, color=MUTED).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Combobox(settings, textvariable=self.start_year, state="readonly", width=8,
                     values=[str(year) for year in range(2026, 2031)]).grid(
            row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Combobox(settings, textvariable=self.course_limit, state="readonly", width=8,
                     values=[str(value) for value in range(1, 7)]).grid(
            row=1, column=1, sticky="ew", padx=(8, 0), pady=(6, 0))

        self._text(panel, "PROJECT DATA", 9, bold=True, color=MUTED).pack(
            anchor="w", pady=(12, 7))
        for description, file_name in (("Prerequisites", "prereq courses.txt"),
                                       ("Class schedule", "data/course_schedule.csv")):
            ready = (PROJECT_FOLDER / file_name).is_file()
            color = "#287047" if ready else "#AC422E"
            self._text(panel, f"{'✓' if ready else '!'}  {description}: {'Ready' if ready else 'Missing'}",
                       9, color=color).pack(anchor="w", pady=3)

    def _build_results(self, parent):
        # Provides independent tabs for generating and checking class plans.
        self._text(parent, "Planning results", 16, bold=True).pack(
            anchor="w", padx=22, pady=(20, 5))
        tk.Label(parent, textvariable=self.summary, bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 10), anchor="w", justify="left",
                 wraplength=530).pack(fill="x", padx=22, pady=(0, 8))

        buttons = tk.Frame(parent, bg=WHITE)
        buttons.pack(side="bottom", fill="x", padx=22, pady=16)
        self.open_button = self._button(buttons, "Open Excel Plan", self._open_excel, primary=True)
        self.open_button.pack(side="left")
        self.open_button.configure(state="disabled")
        self.print_button = self._button(buttons, "Print Plan", self._print_excel)
        self.print_button.pack(side="left", padx=8)
        self.print_button.configure(state="disabled")
        self.graph_button = self._button(buttons, "View Prerequisite Graph", self._graph)
        self.graph_button.pack(side="left")
        self.graph_button.configure(state="disabled")

        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True, padx=22, pady=(2, 0))
        plan_frame = tk.Frame(notebook, bg=WHITE)
        check_frame = tk.Frame(notebook, bg=WHITE)
        notebook.add(plan_frame, text="Recommended Plan")
        notebook.add(check_frame, text="Check Existing Excel Plan")

        table_container = tk.Frame(plan_frame, bg=WHITE)
        table_container.pack(fill="both", expand=True, padx=6, pady=8)
        cols = ("term", "code", "title", "credits")
        self.plan_table = ttk.Treeview(table_container, columns=cols, show="headings", style="Course.Treeview")
        for col, title, width in (("term", "Semester", 100), ("code", "Course", 102),
                                  ("title", "Course Title", 195), ("credits", "Credits", 60)):
            self.plan_table.heading(col, text=title, anchor="w")
            self.plan_table.column(col, width=width, minwidth=55, stretch=(col == "title"))
        scroll = ttk.Scrollbar(table_container, orient="vertical", command=self.plan_table.yview)
        self.plan_table.configure(yscrollcommand=scroll.set)
        self.plan_table.pack(side="left", expand=True, fill="both")
        scroll.pack(side="right", fill="y")
        self.plan_table.tag_configure("pending", foreground="#A15F13")

        self._text(check_frame, "Upload an Excel plan to check prerequisite order.",
                   10, color=MUTED).pack(fill="x", padx=16, pady=(20, 12))
        self._button(check_frame, "Choose and Check Excel File...", self._check_excel).pack(
            anchor="w", padx=16)
        self.check_result = tk.Text(check_frame, wrap="word", font=("Segoe UI", 10),
                                    bg=WHITE, fg=TEXT, padx=12, pady=12, relief="flat")
        self.check_result.pack(fill="both", expand=True, padx=16, pady=14)
        self.check_result.insert("1.0", "No Excel file checked yet.")
        self.check_result.configure(state="disabled")

    def _choose_program(self):
        # Opens a small search dialog containing the supported degree program.
        dialog = tk.Toplevel(self)
        dialog.title("Choose Degree Program")
        dialog.geometry("440x210")
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg=WHITE)
        tk.Label(dialog, text="Search degree program", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(18, 8))
        search = tk.StringVar()
        field = ttk.Entry(dialog, textvariable=search)
        field.pack(fill="x", padx=20)
        option = tk.Listbox(dialog, font=("Segoe UI", 10), height=2)
        option.pack(fill="x", padx=20, pady=12)
        option.insert("end", PROGRAM_NAME)

        def filter_program(*_args):
            # Filters the supported degree list using the search text.
            option.delete(0, "end")
            if search.get().lower() in PROGRAM_NAME.lower():
                option.insert("end", PROGRAM_NAME)

        search.trace_add("write", filter_program)
        def select_program():
            # Stores the selected degree and closes the search dialog.
            if option.size():
                self.program.set(PROGRAM_NAME)
                dialog.destroy()
        self._button(dialog, "Select program", select_program, primary=True).pack(
            anchor="e", padx=20)
        field.focus_set()

    def _choose_pdf(self):
        # Opens the standard Windows file picker for a DegreeWorks PDF.
        selected = filedialog.askopenfilename(title="Choose DegreeWorks PDF",
                                              filetypes=[("PDF files", "*.pdf")])
        if selected:
            self.pdf_file.set(selected)

    def _sample_pdf(self):
        # Selects the demonstration PDF bundled with the original project.
        sample = PROJECT_FOLDER / "Sample Input1.pdf"
        if sample.is_file():
            self.pdf_file.set(str(sample))
        else:
            messagebox.showerror("Missing PDF", f"Sample Input1.pdf was not found in {PROJECT_FOLDER}")

    def _generate(self):
        # Generates the plan and Excel file from the three required project inputs.
        pdf = Path(self.pdf_file.get())
        if not pdf.is_file():
            messagebox.showwarning("DegreeWorks PDF", "Choose a DegreeWorks PDF first.")
            return
        try:
            self.status.set("Generating course plan...")
            self.update_idletasks()
            result = generate_plan(PROJECT_FOLDER, pdf, self.start_year.get(), self.course_limit.get())
            excel_file = export_plan(result, PROJECT_FOLDER)
        except Exception as error:
            self.status.set("Could not generate the plan.")
            messagebox.showerror("Planning error", str(error))
            return

        self.generated_result = result
        self.excel_path = excel_file
        for row in self.plan_table.get_children():
            self.plan_table.delete(row)
        for semester, code, title, credits in result["rows"]:
            self.plan_table.insert("", "end", values=(semester, code, title, credits))
        for code, reason in result["unscheduled"]:
            self.plan_table.insert("", "end", values=("Unscheduled", code, reason, ""), tags=("pending",))

        self.summary.set(
            f"{len(result['rows'])} of {len(result['remaining'])} required courses scheduled • "
            f"{result['total_credits']} credits. "
            f"Excel saved to output/{excel_file.name}."
        )
        self.status.set(f"Plan generated: {excel_file}")
        for button in (self.open_button, self.print_button, self.graph_button):
            button.configure(state="normal")
        if result["unscheduled"]:
            messagebox.showinfo("Plan generated", f"Excel created successfully. "
                                f"{len(result['unscheduled'])} course(s) are listed as Unscheduled.")
        else:
            messagebox.showinfo("Plan generated", "Excel created successfully in the output folder.")

    def _open_excel(self):
        # Opens the automatically generated workbook using the default Windows application.
        if self.excel_path and self.excel_path.is_file():
            try:
                os.startfile(str(self.excel_path))
            except OSError as error:
                messagebox.showerror("Open Excel", str(error))

    def _print_excel(self):
        # Starts the Windows print action after confirming the requested workbook.
        if not self.excel_path or not self.excel_path.is_file():
            return
        if messagebox.askyesno("Print plan", "Print the generated Excel plan?"):
            try:
                os.startfile(str(self.excel_path), "print")
            except OSError as error:
                messagebox.showerror("Print plan", str(error))

    def _check_excel(self):
        # Checks an existing workbook without affecting the generated class plan.
        selected = filedialog.askopenfilename(title="Select existing Excel class plan",
                                              filetypes=[("Excel workbook", "*.xlsx")])
        if not selected:
            return
        try:
            courses, errors = check_excel_plan(selected, PROJECT_FOLDER)
        except Exception as error:
            messagebox.showerror("Excel check", str(error))
            return
        lines = [f"Checked {len(courses)} scheduled course entries.", ""]
        if errors:
            lines.append(f"Found {len(errors)} prerequisite / duplicate issue(s):")
            lines.extend(f"• {error}" for error in errors)
        else:
            lines.append("No prerequisite order problems were found among the courses in this plan.")
        lines.append("")
        lines.append("Prerequisites not listed in the workbook are assumed to be completed earlier.")
        self.check_result.configure(state="normal")
        self.check_result.delete("1.0", "end")
        self.check_result.insert("1.0", "\n".join(lines))
        self.check_result.configure(state="disabled")
        self.status.set("Existing Excel plan checked.")

    def _graph(self):
        # Displays prerequisite relationships using the existing NetworkX graph.
        if not self.generated_result:
            return
        try:
            import matplotlib.pyplot as plt
            import networkx as nx
            graph = self.generated_result["graph"]
            subset = set(self.generated_result["remaining"])
            for code in self.generated_result["remaining"]:
                if code in graph:
                    subset.update(nx.ancestors(graph, code))
            visible_graph = graph.subgraph(subset).copy()
            visible_graph.add_nodes_from(self.generated_result["remaining"])
            plt.figure(figsize=(11, 7))
            positions = nx.spring_layout(visible_graph, seed=7)
            nx.draw_networkx(visible_graph, positions, with_labels=True,
                             node_size=1900, font_size=8, arrows=True)
            plt.title("Prerequisite Graph")
            plt.axis("off")
            plt.tight_layout()
            plt.show()
        except Exception as error:
            messagebox.showerror("Prerequisite graph", str(error))


if __name__ == "__main__":
    SmartPlanningApp().mainloop()
