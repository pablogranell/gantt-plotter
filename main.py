from src.utils.gantt_utils import load_tasks, group_tasks_by_group, plot_gantt

def main():
    
    file_path = 'Gantt.xlsx'
    output_path = None
    
    # Excel file specs
    sheet_name = 'Gantt'
    header = 0
    nrows = 52
    skiprows = None 

    tasks = load_tasks(file_path, sheet_name, header, nrows, skiprows)
    grouped_tasks = group_tasks_by_group(tasks)
    plot_gantt(grouped_tasks, output_path)

if __name__ == "__main__":
    main()