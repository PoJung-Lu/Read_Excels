"""
GUI Launcher for Excel Data Processing Application

Directory Structure:
Project_A/
├── ExcelProcessor.exe  (this executable)
├── config.yaml         (configuration file)
└── Data/
    ├── 消防機關救災能量/
    └── 高科技產業/
"""

import logging
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import yaml

from Read_excels_as_one import (
    firefighter_training_survey_main,
    high_tech_industry_rescue_equipment_main,
    high_tech_industry_chems_main,
)


def get_executable_dir():
    """Get the directory where the executable is located"""
    if getattr(sys, "frozen", False):
        # Running as compiled executable
        return Path(sys.executable).parent
    else:
        # Running as script
        return Path(__file__).parent


class TextHandler(logging.Handler):
    """Custom logging handler to redirect logs to tkinter text widget"""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        try:
            self.text_widget.configure(state="normal")
            self.text_widget.insert(tk.END, msg + "\n")
            self.text_widget.configure(state="disabled")
            self.text_widget.see(tk.END)
            self.text_widget.update_idletasks()  # Force GUI update
        except Exception:
            pass  # Avoid errors if widget is destroyed


class TextRedirector:
    """Redirect stdout/stderr to tkinter text widget for print statements"""

    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        if message.strip():  # Only write non-empty messages
            try:
                self.text_widget.configure(state="normal")
                self.text_widget.insert(tk.END, message)
                self.text_widget.configure(state="disabled")
                self.text_widget.see(tk.END)
                self.text_widget.update_idletasks()
            except Exception:
                pass

    def flush(self):
        pass  # Required for file-like object interface


class ConfigManager:
    """Manages configuration loading and saving"""

    def __init__(self, config_path):
        self.config_path = config_path
        self.config = self.load_config()

    def load_config(self):
        """Load configuration from YAML file"""
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return yaml.safe_load(f)
            except Exception as e:
                messagebox.showwarning(
                    "Config Warning",
                    f"Could not load config file: {e}\nUsing defaults.",
                )
        return self.get_default_config()

    def get_default_config(self):
        """Return default configuration"""
        return {
            "firefighter": {
                "base": "./Data/消防機關救災能量",
                "output": "/../Output",
                "enabled": True,
            },
            "industry": {
                "base": "./Data/科技廠救災能量",
                "output": "/../Output",
                "enabled": True,
            },
            "general": {"auto_run": False, "show_console": True},
        }

    def save_config(self):
        """Save current configuration to YAML file"""
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                yaml.dump(self.config, f, allow_unicode=True, default_flow_style=False)
            return True
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save config: {e}")
            return False


class ExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Processing Tool")
        self.root.geometry("900x700")

        # Get executable directory
        self.exe_dir = get_executable_dir()

        # Load configuration
        self.config_manager = ConfigManager(self.exe_dir / "config.yaml")
        self.config = self.config_manager.config

        # Save original stdout/stderr
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr

        # Configure logging
        self.setup_logging()

        # Create UI
        self.create_ui()

        # Show executable location
        self.log_message(f"Working Directory: {self.exe_dir}", "INFO")

    def setup_logging(self):
        """Configure logging"""
        logging.getLogger().handlers.clear()
        logging.basicConfig(
            level=logging.INFO,
            format="%(levelname)s: %(message)s",
            handlers=[],
        )

    def create_ui(self):
        """Create the user interface"""
        # Create main frame with scrollbar
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title = ttk.Label(
            main_frame,
            text="Excel Data Processing Application",
            font=("Arial", 16, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, pady=10)

        # Working directory display
        dir_frame = ttk.Frame(main_frame)
        dir_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        ttk.Label(dir_frame, text="Working Directory:", font=("Arial", 9, "bold")).pack(
            side=tk.LEFT
        )
        ttk.Label(dir_frame, text=str(self.exe_dir), font=("Arial", 9)).pack(
            side=tk.LEFT, padx=5
        )

        # Firefighter Survey Section
        ff_frame = ttk.LabelFrame(
            main_frame, text="Firefighter Training Survey", padding="10"
        )
        ff_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.ff_enabled_var = tk.BooleanVar(
            value=self.config["firefighter"].get("enabled", True)
        )
        ttk.Checkbutton(ff_frame, text="Enable", variable=self.ff_enabled_var).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )

        ttk.Label(ff_frame, text="Base Directory:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.ff_base_var = tk.StringVar(value=self.config["firefighter"].get("base"))
        ttk.Entry(ff_frame, textvariable=self.ff_base_var, width=60).grid(
            row=1, column=1, padx=5, sticky=(tk.W, tk.E)
        )
        ttk.Button(ff_frame, text="Browse", command=self.browse_ff_base).grid(
            row=1, column=2
        )

        ttk.Label(ff_frame, text="Output (relative):").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.ff_output_var = tk.StringVar(
            value=self.config["firefighter"].get("output")
        )
        ttk.Entry(ff_frame, textvariable=self.ff_output_var, width=60).grid(
            row=2, column=1, padx=5, sticky=(tk.W, tk.E)
        )

        # Industry Analysis Section
        ind_frame = ttk.LabelFrame(
            main_frame,
            text="High-Tech Industry (Chemical Storage + Rescue Equipment)",
            padding="10",
        )
        ind_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.ind_enabled_var = tk.BooleanVar(
            value=self.config["industry"].get("enabled", True)
        )
        ttk.Checkbutton(ind_frame, text="Enable", variable=self.ind_enabled_var).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )

        ttk.Label(ind_frame, text="Base Directory:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.ind_base_var = tk.StringVar(value=self.config["industry"].get("base"))
        ttk.Entry(ind_frame, textvariable=self.ind_base_var, width=60).grid(
            row=1, column=1, padx=5, sticky=(tk.W, tk.E)
        )
        ttk.Button(ind_frame, text="Browse", command=self.browse_ind_base).grid(
            row=1, column=2
        )

        ttk.Label(ind_frame, text="Output (relative):").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.ind_output_var = tk.StringVar(value=self.config["industry"].get("output"))
        ttk.Entry(ind_frame, textvariable=self.ind_output_var, width=60).grid(
            row=2, column=1, padx=5, sticky=(tk.W, tk.E)
        )

        ff_frame.columnconfigure(1, weight=1)
        ind_frame.columnconfigure(1, weight=1)

        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10)

        ttk.Button(
            button_frame,
            text="Run Firefighter Analysis",
            command=self.run_firefighter_analysis,
            width=25,
        ).grid(row=0, column=0, padx=5)

        ttk.Button(
            button_frame,
            text="Run Industry Analysis",
            command=self.run_industry_analysis,
            width=25,
        ).grid(row=0, column=1, padx=5)

        ttk.Button(
            button_frame, text="Run Both", command=self.run_both_analyses, width=25
        ).grid(row=0, column=2, padx=5)

        # Verbose logging checkbox
        verbose_frame = ttk.Frame(main_frame)
        verbose_frame.grid(row=5, column=0, columnspan=3, pady=5)

        self.verbose_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            verbose_frame,
            text="Show detailed logs (including all print statements)",
            variable=self.verbose_var,
        ).pack()

        # Config buttons
        config_button_frame = ttk.Frame(main_frame)
        config_button_frame.grid(row=6, column=0, columnspan=3, pady=5)

        ttk.Button(
            config_button_frame,
            text="Save Configuration",
            command=self.save_configuration,
        ).grid(row=0, column=0, padx=5)

        ttk.Button(
            config_button_frame, text="Reset to Defaults", command=self.reset_defaults
        ).grid(row=0, column=1, padx=5)

        # Log output area
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=20, state="disabled", wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)

    def log_message(self, message, level="INFO"):
        """Add a message to the log text widget"""
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"{level}: {message}\n")
        self.log_text.configure(state="disabled")
        self.log_text.see(tk.END)
        self.root.update()

    def add_text_handler(self):
        """Add text widget handler for logging"""
        # Remove existing handlers
        logger = logging.getLogger()
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)

        # Add new text handler
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        logger.addHandler(text_handler)

        # If verbose mode is enabled, also redirect stdout/stderr
        if self.verbose_var.get():
            sys.stdout = TextRedirector(self.log_text)
            sys.stderr = TextRedirector(self.log_text)
        else:
            # Restore original stdout/stderr when verbose is disabled
            sys.stdout = self.original_stdout
            sys.stderr = self.original_stderr

    def resolve_path(self, path_str):
        """Resolve path relative to executable directory"""
        path = Path(path_str)
        if not path.is_absolute():
            path = self.exe_dir / path
        return str(path.resolve())

    def browse_ff_base(self):
        """Browse for firefighter base directory"""
        directory = filedialog.askdirectory(
            title="Select Firefighter Data Directory", initialdir=self.exe_dir
        )
        if directory:
            # Convert to relative path if under exe_dir
            try:
                rel_path = Path(directory).relative_to(self.exe_dir)
                self.ff_base_var.set(f"./{rel_path}")
            except ValueError:
                self.ff_base_var.set(directory)

    def browse_ind_base(self):
        """Browse for industry base directory"""
        directory = filedialog.askdirectory(
            title="Select Industry Data Directory", initialdir=self.exe_dir
        )
        if directory:
            # Convert to relative path if under exe_dir
            try:
                rel_path = Path(directory).relative_to(self.exe_dir)
                self.ind_base_var.set(f"./{rel_path}")
            except ValueError:
                self.ind_base_var.set(directory)

    def save_configuration(self):
        """Save current settings to config file"""
        self.config["firefighter"]["base"] = self.ff_base_var.get()
        self.config["firefighter"]["output"] = self.ff_output_var.get()
        self.config["firefighter"]["enabled"] = self.ff_enabled_var.get()

        self.config["industry"]["base"] = self.ind_base_var.get()
        self.config["industry"]["output"] = self.ind_output_var.get()
        self.config["industry"]["enabled"] = self.ind_enabled_var.get()

        if self.config_manager.save_config():
            messagebox.showinfo("Success", "Configuration saved successfully!")
            self.log_message("Configuration saved to config.yaml", "INFO")

    def reset_defaults(self):
        """Reset configuration to defaults"""
        if messagebox.askyesno(
            "Confirm Reset", "Reset all settings to default values?"
        ):
            self.config = self.config_manager.get_default_config()
            self.ff_base_var.set(self.config["firefighter"]["base"])
            self.ff_output_var.set(self.config["firefighter"]["output"])
            self.ff_enabled_var.set(self.config["firefighter"]["enabled"])
            self.ind_base_var.set(self.config["industry"]["base"])
            self.ind_output_var.set(self.config["industry"]["output"])
            self.ind_enabled_var.set(self.config["industry"]["enabled"])
            self.log_message("Configuration reset to defaults", "INFO")

    def run_firefighter_analysis(self):
        """Run firefighter training survey analysis"""
        if not self.ff_enabled_var.get():
            messagebox.showinfo("Skipped", "Firefighter analysis is disabled")
            return

        self.log_text.configure(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state="disabled")

        self.add_text_handler()

        base = self.resolve_path(self.ff_base_var.get())
        output = self.ff_output_var.get()

        if not Path(base).exists():
            messagebox.showerror("Error", f"Base directory does not exist:\n{base}")
            return

        try:
            self.log_message(f"Starting firefighter analysis...", "INFO")
            self.log_message(f"Base: {base}", "INFO")
            self.log_message(f"Output: {output}\n", "INFO")

            firefighter_training_survey_main(base=base, out_rel=output)

            self.log_message("\n✓ Firefighter analysis completed!", "INFO")
            messagebox.showinfo(
                "Success", "Firefighter analysis completed successfully!"
            )
        except Exception as e:
            logging.error(f"Error during firefighter analysis: {str(e)}")
            messagebox.showerror("Error", f"Analysis failed:\n{str(e)}")

    def run_industry_analysis(self):
        """Run high-tech industry rescue equipment analysis"""
        if not self.ind_enabled_var.get():
            messagebox.showinfo("Skipped", "Industry analysis is disabled")
            return

        self.log_text.configure(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state="disabled")

        self.add_text_handler()

        base = self.resolve_path(self.ind_base_var.get())
        output = self.ind_output_var.get()

        if not Path(base).exists():
            messagebox.showerror("Error", f"Base directory does not exist:\n{base}")
            return

        try:
            self.log_message(f"Starting industry analysis...", "INFO")
            self.log_message(f"Base: {base}", "INFO")
            self.log_message(f"Output: {output}\n", "INFO")

            # Run chemical storage analysis first
            self.log_message("Step 1: Chemical Storage Analysis", "INFO")
            high_tech_industry_chems_main(base=base, out_rel=output)

            # Then run rescue equipment analysis
            self.log_message("\nStep 2: Rescue Equipment Analysis", "INFO")
            high_tech_industry_rescue_equipment_main(base=base)

            self.log_message("\n✓ Industry analysis completed!", "INFO")
            messagebox.showinfo("Success", "Industry analysis completed successfully!")
        except Exception as e:
            logging.error(f"Error during industry analysis: {str(e)}")
            messagebox.showerror("Error", f"Analysis failed:\n{str(e)}")

    def run_both_analyses(self):
        """Run both analyses sequentially"""
        self.log_text.configure(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state="disabled")

        self.add_text_handler()

        results = []

        # Run firefighter analysis
        if self.ff_enabled_var.get():
            try:
                base = self.resolve_path(self.ff_base_var.get())
                output = self.ff_output_var.get()

                if Path(base).exists():
                    self.log_message("=" * 60, "INFO")
                    self.log_message("FIREFIGHTER ANALYSIS", "INFO")
                    self.log_message("=" * 60, "INFO")
                    firefighter_training_survey_main(base=base, out_rel=output)
                    results.append("✓ Firefighter analysis completed")
                else:
                    results.append(f"✗ Firefighter: Directory not found: {base}")
            except Exception as e:
                results.append(f"✗ Firefighter analysis failed: {str(e)}")
                logging.error(f"Firefighter analysis error: {str(e)}")

        # Run industry analysis
        if self.ind_enabled_var.get():
            try:
                base = self.resolve_path(self.ind_base_var.get())
                output = self.ind_output_var.get()

                if Path(base).exists():
                    self.log_message("\n" + "=" * 60, "INFO")
                    self.log_message("INDUSTRY ANALYSIS", "INFO")
                    self.log_message("=" * 60, "INFO")
                    self.log_message("Step 1: Chemical Storage Analysis", "INFO")
                    high_tech_industry_chems_main(base=base, out_rel=output)
                    self.log_message("\nStep 2: Rescue Equipment Analysis", "INFO")
                    
                    results.append("✓ Industry analysis completed")
                else:
                    results.append(f"✗ Industry: Directory not found: {base}")
            except Exception as e:
                results.append(f"✗ Industry analysis failed: {str(e)}")
                logging.error(f"Industry analysis error: {str(e)}")

        # Show summary
        summary = "\n".join(results)
        self.log_message("\n" + "=" * 60, "INFO")
        self.log_message("SUMMARY", "INFO")
        self.log_message("=" * 60, "INFO")
        self.log_message(summary, "INFO")

        messagebox.showinfo("Batch Processing Complete", summary)


def main():
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
