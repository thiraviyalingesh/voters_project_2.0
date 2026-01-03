"""
Electoral Roll Voter Analytics Dashboard
Dashboard with constituency selection, gender stats, multiple charts, and PDF export
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime


class VoterAnalyticsDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Voter Analytics Dashboard")
        self.root.geometry("1300x900")
        self.root.resizable(True, True)
        self.root.configure(bg='#f0f0f0')

        self.df = None
        self.excel_path = None
        self.constituencies = []

        # Default age ranges
        self.age_ranges = [
            (18, 25, "18-25"),
            (26, 35, "26-35"),
            (36, 45, "36-45"),
            (46, 55, "46-55"),
            (56, 65, "56-65"),
            (66, 75, "66-75"),
            (76, 150, "76+")
        ]

        # Store current stats for PDF
        self.current_stats = {}

        self.create_widgets()

    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))

        title_label = tk.Label(title_frame, text="Voter Analytics Dashboard",
                               font=('Helvetica', 20, 'bold'), bg='#f0f0f0', fg='#333')
        title_label.pack()

        # Top bar: File selection + Constituency dropdown
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # File selection
        ttk.Label(top_frame, text="Excel File:").pack(side=tk.LEFT, padx=(0, 5))
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(top_frame, textvariable=self.file_path_var, width=40)
        file_entry.pack(side=tk.LEFT, padx=(0, 5))

        browse_btn = ttk.Button(top_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Constituency dropdown
        ttk.Label(top_frame, text="Constituency:").pack(side=tk.LEFT, padx=(0, 5))
        self.constituency_var = tk.StringVar()
        self.constituency_combo = ttk.Combobox(top_frame, textvariable=self.constituency_var,
                                                state='readonly', width=25)
        self.constituency_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.constituency_combo.bind('<<ComboboxSelected>>', self.on_constituency_change)

        # Age range config
        ttk.Label(top_frame, text="Age Ranges:").pack(side=tk.LEFT, padx=(0, 5))
        self.age_range_var = tk.StringVar(value="18-25,26-35,36-45,46-55,56-65,66-75,76+")
        age_entry = ttk.Entry(top_frame, textvariable=self.age_range_var, width=30)
        age_entry.pack(side=tk.LEFT, padx=(0, 5))

        apply_btn = ttk.Button(top_frame, text="Apply", command=self.apply_and_refresh)
        apply_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Export PDF button
        export_btn = ttk.Button(top_frame, text="Export PDF", command=self.export_pdf)
        export_btn.pack(side=tk.LEFT, padx=(5, 0))

        # Stats frame (Total + Percentages only)
        stats_frame = ttk.Frame(main_frame)
        stats_frame.pack(fill=tk.X, pady=(0, 10))

        # Create stat boxes - only Total and percentages
        self.create_stat_box(stats_frame, "total_box", "Total Voters", "#2196F3", 0)
        self.create_stat_box(stats_frame, "male_pct_box", "Male %", "#1565C0", 1)
        self.create_stat_box(stats_frame, "female_pct_box", "Female %", "#E91E63", 2)
        self.create_stat_box(stats_frame, "other_pct_box", "Third Gender %", "#FF5722", 3)

        # Charts container
        charts_frame = ttk.Frame(main_frame)
        charts_frame.pack(fill=tk.BOTH, expand=True)

        # Top row charts (Gender Pie + Religion Pie + Stacked Bar)
        top_charts_frame = ttk.Frame(charts_frame)
        top_charts_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        # Gender Pie chart frame (left)
        pie_frame = ttk.LabelFrame(top_charts_frame, text="Gender Distribution", padding="5")
        pie_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.pie_fig = Figure(figsize=(3.5, 3.5), dpi=100)
        self.pie_ax = self.pie_fig.add_subplot(111)
        self.pie_canvas = FigureCanvasTkAgg(self.pie_fig, master=pie_frame)
        self.pie_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Religion Pie chart frame (middle)
        religion_frame = ttk.LabelFrame(top_charts_frame, text="Religion Distribution", padding="5")
        religion_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.religion_fig = Figure(figsize=(3.5, 3.5), dpi=100)
        self.religion_ax = self.religion_fig.add_subplot(111)
        self.religion_canvas = FigureCanvasTkAgg(self.religion_fig, master=religion_frame)
        self.religion_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Stacked bar chart frame (right)
        stacked_frame = ttk.LabelFrame(top_charts_frame, text="Gender by Age Group", padding="5")
        stacked_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.stacked_fig = Figure(figsize=(6, 3.5), dpi=100)
        self.stacked_ax = self.stacked_fig.add_subplot(111)
        self.stacked_canvas = FigureCanvasTkAgg(self.stacked_fig, master=stacked_frame)
        self.stacked_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Bottom chart (Age distribution horizontal bar)
        bar_frame = ttk.LabelFrame(charts_frame, text="Age-wise Voter Distribution (Descending Order)", padding="5")
        bar_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        self.bar_fig = Figure(figsize=(10, 3.5), dpi=100)
        self.bar_ax = self.bar_fig.add_subplot(111)
        self.bar_canvas = FigureCanvasTkAgg(self.bar_fig, master=bar_frame)
        self.bar_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Status bar
        self.status_var = tk.StringVar(value="Ready. Please select an Excel file.")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(10, 0))

    def create_stat_box(self, parent, name, title, color, column):
        """Create a statistics display box."""
        frame = tk.Frame(parent, bg=color, padx=15, pady=10)
        frame.grid(row=0, column=column, padx=3, sticky='nsew')
        parent.columnconfigure(column, weight=1)

        title_lbl = tk.Label(frame, text=title, font=('Helvetica', 10), bg=color, fg='white')
        title_lbl.pack()

        value_var = tk.StringVar(value="--")
        value_lbl = tk.Label(frame, textvariable=value_var, font=('Helvetica', 18, 'bold'),
                             bg=color, fg='white')
        value_lbl.pack()

        setattr(self, f"{name}_var", value_var)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=Path(__file__).parent
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.load_file()

    def load_file(self):
        file_path = self.file_path_var.get()
        if not file_path or not Path(file_path).exists():
            return

        try:
            self.status_var.set("Loading Excel file...")
            self.root.update()

            self.df = pd.read_excel(file_path)
            self.excel_path = file_path

            # Clean data
            self.clean_data()

            # Get constituencies
            if 'Constituency' in self.df.columns:
                self.constituencies = sorted(self.df['Constituency'].dropna().unique().tolist())
            else:
                # Use filename as constituency
                self.constituencies = [Path(file_path).stem.replace('_excel', '')]

            # Update combobox
            self.constituency_combo['values'] = self.constituencies
            if self.constituencies:
                self.constituency_combo.current(0)
                self.constituency_var.set(self.constituencies[0])

            # Update display
            self.update_dashboard()

            self.status_var.set(f"Loaded {len(self.df):,} records")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
            self.status_var.set("Error loading file.")

    def clean_data(self):
        """Clean and prepare data for analysis."""
        if self.df is None:
            return

        # Convert Age to numeric
        if 'Age' in self.df.columns:
            self.df['Age_Clean'] = pd.to_numeric(self.df['Age'], errors='coerce')

        # Standardize Gender - check multiple possible column names
        gender_col = None
        for col_name in ['Gender', 'gender', 'GENDER', 'Sex', 'sex', 'SEX']:
            if col_name in self.df.columns:
                gender_col = col_name
                break

        if gender_col:
            self.df['Gender_Clean'] = self.df[gender_col].astype(str).str.strip().str.title()
            # Normalize Male variations
            self.df.loc[self.df['Gender_Clean'].isin(['M', 'Male', 'Man', 'Boy']), 'Gender_Clean'] = 'Male'
            # Normalize Female variations
            self.df.loc[self.df['Gender_Clean'].isin(['F', 'Female', 'Woman', 'Girl']), 'Gender_Clean'] = 'Female'
            # Everything else (missing, nan, empty, other) stays as-is and will be counted as Third Gender
            # This includes: 'Nan', 'None', '', 'Other', 'Transgender', etc.

        # Standardize Religion - check multiple possible column names
        religion_col = None
        for col_name in ['Religion', 'religion', 'RELIGION', 'Rel', 'rel', 'REL']:
            if col_name in self.df.columns:
                religion_col = col_name
                break

        if religion_col:
            self.df['Religion_Clean'] = self.df[religion_col].astype(str).str.strip().str.title()
            # Replace 'Nan' string with empty
            self.df.loc[self.df['Religion_Clean'].str.lower() == 'nan', 'Religion_Clean'] = ''
            self.df.loc[self.df['Religion_Clean'] == 'None', 'Religion_Clean'] = ''
            self.df.loc[self.df['Religion_Clean'] == '', 'Religion_Clean'] = 'Unknown'

    def parse_age_ranges(self, range_string):
        """Parse age range string like '18-25,26-35,36+' into list of tuples."""
        ranges = []
        parts = range_string.split(',')

        for part in parts:
            part = part.strip()
            if not part:
                continue

            if '+' in part:
                start = int(part.replace('+', ''))
                ranges.append((start, 150, part))
            elif '-' in part:
                start, end = part.split('-')
                ranges.append((int(start), int(end), part))

        return ranges

    def apply_and_refresh(self):
        try:
            self.age_ranges = self.parse_age_ranges(self.age_range_var.get())
            self.update_dashboard()
            self.status_var.set(f"Applied {len(self.age_ranges)} age ranges.")
        except Exception as e:
            messagebox.showerror("Error", f"Invalid age range format: {e}")

    def on_constituency_change(self, event=None):
        self.update_dashboard()

    def get_filtered_df(self):
        """Get dataframe filtered by selected constituency."""
        if self.df is None:
            return None

        constituency = self.constituency_var.get()
        if 'Constituency' in self.df.columns and constituency:
            return self.df[self.df['Constituency'] == constituency]
        return self.df

    def update_dashboard(self):
        """Update all dashboard elements."""
        df = self.get_filtered_df()
        if df is None or len(df) == 0:
            return

        total = len(df)

        # Gender stats
        male_count = 0
        female_count = 0
        other_count = 0
        if 'Gender_Clean' in df.columns:
            male_count = (df['Gender_Clean'] == 'Male').sum()
            female_count = (df['Gender_Clean'] == 'Female').sum()
            other_count = total - male_count - female_count

        male_pct_raw = (male_count / total * 100) if total > 0 else 0
        female_pct_raw = (female_count / total * 100) if total > 0 else 0
        other_pct_raw = (other_count / total * 100) if total > 0 else 0

        # Round to 2 decimal places and adjust largest to ensure sum = 100.00%
        male_pct = round(male_pct_raw, 2)
        female_pct = round(female_pct_raw, 2)
        other_pct = round(other_pct_raw, 2)

        # Fix rounding error - adjust the largest percentage
        total_pct = male_pct + female_pct + other_pct
        if total_pct != 100.0 and total > 0:
            diff = round(100.0 - total_pct, 2)
            # Adjust the largest percentage
            if male_pct >= female_pct and male_pct >= other_pct:
                male_pct = round(male_pct + diff, 2)
            elif female_pct >= male_pct and female_pct >= other_pct:
                female_pct = round(female_pct + diff, 2)
            else:
                other_pct = round(other_pct + diff, 2)

        # Religion stats
        religion_stats = {}
        if 'Religion_Clean' in df.columns:
            religion_counts = df['Religion_Clean'].value_counts()
            for religion, count in religion_counts.items():
                religion_str = str(religion).strip().lower()
                if religion and religion_str and religion_str != 'nan' and religion_str != 'none' and religion_str != '':
                    religion_stats[religion] = {
                        'count': count,
                        'pct': (count / total * 100) if total > 0 else 0
                    }

        # Store stats for PDF
        self.current_stats = {
            'total': total,
            'male_count': male_count,
            'female_count': female_count,
            'other_count': other_count,
            'male_pct': male_pct,
            'female_pct': female_pct,
            'other_pct': other_pct,
            'constituency': self.constituency_var.get(),
            'religion_stats': religion_stats
        }

        # Update stat boxes (percentages only) - all use 2 decimal places to add up to 100.00%
        self.total_box_var.set(f"{total:,}")
        self.male_pct_box_var.set(f"{male_pct:.2f}%")
        self.female_pct_box_var.set(f"{female_pct:.2f}%")
        self.other_pct_box_var.set(f"{other_pct:.2f}%")

        # Update all charts
        self.update_pie_chart(df, male_count, female_count, other_count)
        self.update_religion_chart(df)
        self.update_stacked_chart(df)
        self.update_bar_chart(df)

    def update_pie_chart(self, df, male_count, female_count, other_count):
        """Update the gender distribution pie chart."""
        self.pie_ax.clear()

        # Prepare data - Other includes all unknown/missing gender (Third Gender)
        # This ensures Male + Female + Third Gender = Total (100%)
        labels = []
        sizes = []
        colors = []
        explode = []

        if male_count > 0:
            labels.append('Male')
            sizes.append(male_count)
            colors.append('#1565C0')
            explode.append(0.02)

        if female_count > 0:
            labels.append('Female')
            sizes.append(female_count)
            colors.append('#E91E63')
            explode.append(0.02)

        if other_count > 0:
            labels.append('Third Gender')
            sizes.append(other_count)
            colors.append('#FF5722')  # Bright orange for high visibility
            explode.append(0.02)

        if sizes:
            # Use 2 decimal places to show small percentages like 0.25%
            wedges, texts, autotexts = self.pie_ax.pie(
                sizes, labels=labels, colors=colors, explode=explode,
                autopct='%1.2f%%', startangle=90,
                textprops={'fontsize': 9, 'fontweight': 'bold'}
            )
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')

        self.pie_ax.set_title(f'Gender Distribution - {self.constituency_var.get()}',
                              fontsize=11, fontweight='bold', pad=10)

        self.pie_fig.tight_layout()
        self.pie_canvas.draw()

    def update_religion_chart(self, df):
        """Update the religion distribution pie chart."""
        self.religion_ax.clear()

        if 'Religion_Clean' not in df.columns:
            self.religion_ax.text(0.5, 0.5, 'No Religion Data', ha='center', va='center',
                                   fontsize=12, transform=self.religion_ax.transAxes)
            self.religion_ax.set_title(f'Religion Distribution - {self.constituency_var.get()}',
                                        fontsize=11, fontweight='bold', pad=10)
            self.religion_fig.tight_layout()
            self.religion_canvas.draw()
            return

        # Get religion counts
        religion_counts = df['Religion_Clean'].value_counts()

        # Filter out empty/nan values but keep 'Unknown'
        religion_counts = religion_counts[
            (religion_counts.index != '') &
            (religion_counts.index.str.lower() != 'nan') &
            (religion_counts.index.str.lower() != 'none')
        ]

        if len(religion_counts) == 0:
            self.religion_ax.text(0.5, 0.5, 'No Religion Data', ha='center', va='center',
                                   fontsize=12, transform=self.religion_ax.transAxes)
            self.religion_ax.set_title(f'Religion Distribution - {self.constituency_var.get()}',
                                        fontsize=11, fontweight='bold', pad=10)
            self.religion_fig.tight_layout()
            self.religion_canvas.draw()
            return

        # Prepare data - show only percentages, no counts
        labels = []
        sizes = []
        # Color palette for religions
        religion_colors = ['#4CAF50', '#FF9800', '#9C27B0', '#00BCD4', '#E91E63',
                           '#3F51B5', '#FFEB3B', '#795548', '#607D8B', '#F44336']

        for i, (religion, count) in enumerate(religion_counts.items()):
            labels.append(religion)
            sizes.append(count)

        colors = religion_colors[:len(sizes)]
        explode = [0.02] * len(sizes)

        if sizes:
            wedges, texts, autotexts = self.religion_ax.pie(
                sizes, labels=labels, colors=colors, explode=explode,
                autopct='%1.2f%%', startangle=90,
                textprops={'fontsize': 8, 'fontweight': 'bold'}
            )
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(7)

        self.religion_ax.set_title(f'Religion Distribution - {self.constituency_var.get()}',
                                    fontsize=11, fontweight='bold', pad=10)

        self.religion_fig.tight_layout()
        self.religion_canvas.draw()

    def update_stacked_chart(self, df):
        """Update the stacked bar chart showing gender by age group."""
        self.stacked_ax.clear()

        if 'Age_Clean' not in df.columns or 'Gender_Clean' not in df.columns:
            self.stacked_canvas.draw()
            return

        # Calculate male/female counts for each age range
        age_labels = []
        male_counts = []
        female_counts = []
        other_counts = []

        for start, end, label in self.age_ranges:
            mask = (df['Age_Clean'] >= start) & (df['Age_Clean'] <= end)
            age_df = df[mask]

            age_labels.append(label)
            male_counts.append((age_df['Gender_Clean'] == 'Male').sum())
            female_counts.append((age_df['Gender_Clean'] == 'Female').sum())
            other_counts.append(len(age_df) - male_counts[-1] - female_counts[-1])

        x = range(len(age_labels))
        width = 0.6

        # Create stacked bars
        bars1 = self.stacked_ax.bar(x, male_counts, width, label='Male', color='#1565C0')
        bars2 = self.stacked_ax.bar(x, female_counts, width, bottom=male_counts, label='Female', color='#E91E63')

        # Add Third Gender if exists
        if sum(other_counts) > 0:
            bottom = [m + f for m, f in zip(male_counts, female_counts)]
            bars3 = self.stacked_ax.bar(x, other_counts, width, bottom=bottom, label='Third Gender', color='#FF5722')

        # Add value labels on bars
        for i, (m, f) in enumerate(zip(male_counts, female_counts)):
            total = m + f + other_counts[i]
            if total > 0:
                self.stacked_ax.annotate(f'{total:,}',
                                         xy=(i, total),
                                         xytext=(0, 3),
                                         textcoords="offset points",
                                         ha='center', va='bottom',
                                         fontsize=8, fontweight='bold')

        self.stacked_ax.set_xlabel('Age Range', fontsize=10, fontweight='bold')
        self.stacked_ax.set_ylabel('Number of Voters', fontsize=10, fontweight='bold')
        self.stacked_ax.set_title(f'Gender by Age Group - {self.constituency_var.get()}',
                                  fontsize=11, fontweight='bold', pad=10)
        self.stacked_ax.set_xticks(x)
        self.stacked_ax.set_xticklabels(age_labels, fontsize=9)
        self.stacked_ax.legend(loc='upper right', fontsize=9)

        # Style
        self.stacked_ax.spines['top'].set_visible(False)
        self.stacked_ax.spines['right'].set_visible(False)

        self.stacked_fig.tight_layout()
        self.stacked_canvas.draw()

    def update_bar_chart(self, df):
        """Update the age distribution horizontal bar chart."""
        self.bar_ax.clear()

        if 'Age_Clean' not in df.columns:
            self.bar_canvas.draw()
            return

        total = len(df)
        age_data = []

        for start, end, label in self.age_ranges:
            mask = (df['Age_Clean'] >= start) & (df['Age_Clean'] <= end)
            count = mask.sum()
            age_data.append((label, count))

        # Sort by count descending (but reverse for display so highest is at top)
        age_data.sort(key=lambda x: x[1], reverse=False)

        labels = [x[0] for x in age_data]
        counts = [x[1] for x in age_data]

        # Create horizontal bar chart
        colors = plt.cm.Blues([0.3 + 0.5 * (i / len(labels)) for i in range(len(labels))])
        bars = self.bar_ax.barh(labels, counts, color=colors, edgecolor='#333', linewidth=0.5)

        # Add value labels on bars
        for bar, count in zip(bars, counts):
            width = bar.get_width()
            pct = (count / total * 100) if total > 0 else 0
            self.bar_ax.annotate(f'{count:,} ({pct:.1f}%)',
                               xy=(width, bar.get_y() + bar.get_height() / 2),
                               xytext=(5, 0),
                               textcoords="offset points",
                               ha='left', va='center',
                               fontsize=9, fontweight='bold')

        self.bar_ax.set_ylabel('Age Range', fontsize=10, fontweight='bold')
        self.bar_ax.set_xlabel('Number of Voters', fontsize=10, fontweight='bold')
        self.bar_ax.set_title(f'Age Distribution - {self.constituency_var.get()}',
                             fontsize=11, fontweight='bold', pad=10)

        # Style
        self.bar_ax.spines['top'].set_visible(False)
        self.bar_ax.spines['right'].set_visible(False)

        self.bar_fig.tight_layout()
        self.bar_canvas.draw()

    def export_pdf(self):
        """Export all charts and stats to a PDF file."""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first.")
            return

        # Ask for save location
        default_name = f"Voter_Analytics_{self.constituency_var.get()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        file_path = filedialog.asksaveasfilename(
            title="Save PDF Report",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=default_name,
            initialdir=Path(__file__).parent
        )

        if not file_path:
            return

        try:
            self.status_var.set("Generating PDF report...")
            self.root.update()

            with PdfPages(file_path) as pdf:
                # Page 1: Summary and Pie Chart
                fig1 = Figure(figsize=(11, 8.5), dpi=100)

                # Title
                fig1.suptitle(f"Voter Analytics Report - {self.constituency_var.get()}",
                             fontsize=16, fontweight='bold', y=0.98)

                # Add date
                fig1.text(0.5, 0.94, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                         ha='center', fontsize=10)

                # Summary stats as text - all percentages use 2 decimals to add up to 100.00%
                stats_text = f"""
Summary Statistics
══════════════════════════════════════
Total Voters:     {self.current_stats['total']:,}
Male Voters:      {self.current_stats['male_count']:,} ({self.current_stats['male_pct']:.2f}%)
Female Voters:    {self.current_stats['female_count']:,} ({self.current_stats['female_pct']:.2f}%)
Third Gender:     {self.current_stats['other_count']:,} ({self.current_stats['other_pct']:.2f}%)
══════════════════════════════════════
                """

                # Add religion stats if available
                religion_stats = self.current_stats.get('religion_stats', {})
                if religion_stats:
                    stats_text += """
Religion Distribution
══════════════════════════════════════
"""
                    for religion, data in religion_stats.items():
                        stats_text += f"{religion:15} {data['count']:>8,} ({data['pct']:.2f}%)\n"
                    stats_text += "══════════════════════════════════════"

                fig1.text(0.1, 0.80, stats_text, fontsize=11, family='monospace',
                         verticalalignment='top')

                # Pie chart - percentages only (no counts)
                ax1 = fig1.add_axes([0.55, 0.45, 0.4, 0.4])
                labels = []
                sizes = []
                colors = []

                if self.current_stats['male_count'] > 0:
                    labels.append("Male")
                    sizes.append(self.current_stats['male_count'])
                    colors.append('#1565C0')

                if self.current_stats['female_count'] > 0:
                    labels.append("Female")
                    sizes.append(self.current_stats['female_count'])
                    colors.append('#E91E63')

                if self.current_stats['other_count'] > 0:
                    labels.append("Third Gender")
                    sizes.append(self.current_stats['other_count'])
                    colors.append('#FF5722')  # Bright orange for high visibility

                if sizes:
                    ax1.pie(sizes, labels=labels, colors=colors, autopct='%1.2f%%',
                           startangle=90, textprops={'fontsize': 9})
                ax1.set_title('Gender Distribution', fontsize=12, fontweight='bold')

                # Age-wise table
                df = self.get_filtered_df()
                if 'Age_Clean' in df.columns:
                    table_data = [['Age Range', 'Total', 'Male', 'Female', '%']]
                    for start, end, label in self.age_ranges:
                        mask = (df['Age_Clean'] >= start) & (df['Age_Clean'] <= end)
                        age_df = df[mask]
                        total = len(age_df)
                        male = (age_df['Gender_Clean'] == 'Male').sum() if 'Gender_Clean' in age_df.columns else 0
                        female = (age_df['Gender_Clean'] == 'Female').sum() if 'Gender_Clean' in age_df.columns else 0
                        pct = (total / len(df) * 100) if len(df) > 0 else 0
                        table_data.append([label, f'{total:,}', f'{male:,}', f'{female:,}', f'{pct:.1f}%'])

                    ax_table = fig1.add_axes([0.1, 0.08, 0.8, 0.3])
                    ax_table.axis('off')
                    table = ax_table.table(cellText=table_data, loc='center', cellLoc='center',
                                          colWidths=[0.2, 0.2, 0.2, 0.2, 0.2])
                    table.auto_set_font_size(False)
                    table.set_fontsize(10)
                    table.scale(1, 1.5)

                    # Style header row
                    for i in range(5):
                        table[(0, i)].set_facecolor('#2196F3')
                        table[(0, i)].set_text_props(color='white', fontweight='bold')

                pdf.savefig(fig1)
                plt.close(fig1)

                # Page 2: Stacked Bar Chart
                fig2 = Figure(figsize=(11, 8.5), dpi=100)
                ax2 = fig2.add_subplot(111)

                if 'Age_Clean' in df.columns and 'Gender_Clean' in df.columns:
                    age_labels = []
                    male_counts = []
                    female_counts = []
                    other_counts = []

                    for start, end, label in self.age_ranges:
                        mask = (df['Age_Clean'] >= start) & (df['Age_Clean'] <= end)
                        age_df = df[mask]
                        age_labels.append(label)
                        m = (age_df['Gender_Clean'] == 'Male').sum()
                        f = (age_df['Gender_Clean'] == 'Female').sum()
                        male_counts.append(m)
                        female_counts.append(f)
                        other_counts.append(len(age_df) - m - f)

                    x = range(len(age_labels))
                    width = 0.6

                    ax2.bar(x, male_counts, width, label='Male', color='#1565C0')
                    ax2.bar(x, female_counts, width, bottom=male_counts, label='Female', color='#E91E63')

                    # Add Third Gender if exists
                    if sum(other_counts) > 0:
                        bottom = [m + f for m, f in zip(male_counts, female_counts)]
                        ax2.bar(x, other_counts, width, bottom=bottom, label='Third Gender', color='#FF5722')

                    for i, (m, f) in enumerate(zip(male_counts, female_counts)):
                        total_val = m + f + other_counts[i]
                        ax2.annotate(f'{total_val:,}', xy=(i, total_val), xytext=(0, 3),
                                    textcoords="offset points", ha='center', va='bottom',
                                    fontsize=9, fontweight='bold')

                    ax2.set_xlabel('Age Range', fontsize=12, fontweight='bold')
                    ax2.set_ylabel('Number of Voters', fontsize=12, fontweight='bold')
                    ax2.set_title(f'Gender by Age Group - {self.constituency_var.get()}',
                                 fontsize=14, fontweight='bold', pad=15)
                    ax2.set_xticks(x)
                    ax2.set_xticklabels(age_labels)
                    ax2.legend(loc='upper right')
                    ax2.spines['top'].set_visible(False)
                    ax2.spines['right'].set_visible(False)

                fig2.tight_layout()
                pdf.savefig(fig2)
                plt.close(fig2)

                # Page 3: Horizontal Bar Chart
                fig3 = Figure(figsize=(11, 8.5), dpi=100)
                ax3 = fig3.add_subplot(111)

                if 'Age_Clean' in df.columns:
                    total = len(df)
                    age_data = []

                    for start, end, label in self.age_ranges:
                        mask = (df['Age_Clean'] >= start) & (df['Age_Clean'] <= end)
                        count = mask.sum()
                        age_data.append((label, count))

                    age_data.sort(key=lambda x: x[1], reverse=False)
                    labels = [x[0] for x in age_data]
                    counts = [x[1] for x in age_data]

                    colors = plt.cm.Blues([0.3 + 0.5 * (i / len(labels)) for i in range(len(labels))])
                    bars = ax3.barh(labels, counts, color=colors, edgecolor='#333', linewidth=0.5)

                    for bar, count in zip(bars, counts):
                        width = bar.get_width()
                        pct = (count / total * 100) if total > 0 else 0
                        ax3.annotate(f'{count:,} ({pct:.1f}%)',
                                    xy=(width, bar.get_y() + bar.get_height() / 2),
                                    xytext=(5, 0), textcoords="offset points",
                                    ha='left', va='center', fontsize=10, fontweight='bold')

                    ax3.set_ylabel('Age Range', fontsize=12, fontweight='bold')
                    ax3.set_xlabel('Number of Voters', fontsize=12, fontweight='bold')
                    ax3.set_title(f'Age Distribution (Descending) - {self.constituency_var.get()}',
                                 fontsize=14, fontweight='bold', pad=15)
                    ax3.spines['top'].set_visible(False)
                    ax3.spines['right'].set_visible(False)

                fig3.tight_layout()
                pdf.savefig(fig3)
                plt.close(fig3)

                # Page 4: Religion Distribution (if available)
                religion_stats = self.current_stats.get('religion_stats', {})
                if religion_stats:
                    fig4 = Figure(figsize=(11, 8.5), dpi=100)
                    ax4 = fig4.add_subplot(111)

                    labels = []
                    sizes = []
                    religion_colors = ['#4CAF50', '#FF9800', '#9C27B0', '#00BCD4', '#E91E63',
                                       '#3F51B5', '#FFEB3B', '#795548', '#607D8B', '#F44336']

                    for religion, data in religion_stats.items():
                        labels.append(religion)
                        sizes.append(data['count'])

                    colors = religion_colors[:len(sizes)]

                    if sizes:
                        wedges, texts, autotexts = ax4.pie(
                            sizes, labels=labels, colors=colors,
                            autopct='%1.2f%%', startangle=90,
                            textprops={'fontsize': 10, 'fontweight': 'bold'}
                        )
                        for autotext in autotexts:
                            autotext.set_color('white')
                            autotext.set_fontweight('bold')

                    ax4.set_title(f'Religion Distribution - {self.constituency_var.get()}',
                                 fontsize=14, fontweight='bold', pad=15)

                    fig4.tight_layout()
                    pdf.savefig(fig4)
                    plt.close(fig4)

            self.status_var.set(f"PDF exported: {Path(file_path).name}")
            messagebox.showinfo("Success", f"PDF report saved to:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")
            self.status_var.set("Error exporting PDF.")


def main():
    root = tk.Tk()
    app = VoterAnalyticsDashboard(root)
    root.mainloop()


if __name__ == "__main__":
    main()
