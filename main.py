from src.utils.gantt_utils import load_tasks, group_tasks_by_group, plot_gantt
from src.utils.excel_config_gui import show_excel_config

def generate_gantt(config):
    tasks = load_tasks(
        file_path=config["file_path"],
        sheet_name=config["sheet_name"],
        header=config["header"],
        column_mapping=config["column_mapping"]
    )
    
    grouped_tasks = group_tasks_by_group(tasks)
    plot_gantt(grouped_tasks, sheet_name=config["sheet_name"], output_path=None)

def main():
    show_excel_config(on_load=generate_gantt)

if __name__ == "__main__":
    main()