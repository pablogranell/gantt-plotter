import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.patches as mpatches

from matplotlib import rcParams
from config.settings import (
    TITLE, TITLE_SIZE, TITLE_FONT_WEIGHT,
    FONT_FAMILY, FONT_SANS_SERIF, FONT_COLOR,
    LABEL_SIZE, DAY_FONT_SIZE, MONTH_FONT_SIZE, MONTH_FONT_WEIGHT,
    X_LABEL, Y_LABEL, 
    BAR_COLOR, TEAM_BAR_COLORS,
    DATE_FORMAT
)

rcParams['font.family'] = FONT_FAMILY
rcParams['font.sans-serif'] = FONT_SANS_SERIF
rcParams['axes.titlesize'] = TITLE_SIZE
rcParams['axes.labelsize'] = LABEL_SIZE
rcParams['toolbar'] = 'None'  # Quita la barra de herramientas

def load_tasks(file_path, sheet_name, header, nrows, skiprows):
    """
    Loads data from an Excel spreadsheet into a Pandas dataframe.
    """
    try:
        tasks = pd.read_excel(
            file_path, 
            sheet_name=sheet_name, 
            header=header, 
            nrows=nrows, 
            skiprows=skiprows)
        
        tasks.columns = [
            'Fase', 'Tareas', 'Responsable', 'Duración',
            'Fecha Inicio', 'Fecha Fin'
        ]

        tasks['Fecha Inicio'] = pd.to_datetime(tasks['Fecha Inicio'], format=DATE_FORMAT)
        tasks['Fecha Fin'] = pd.to_datetime(tasks['Fecha Fin'], format=DATE_FORMAT)
        tasks.set_index(pd.DatetimeIndex(tasks['Fecha Inicio'].values), inplace=True)
        
        return tasks
    
    except Exception as e:
        print(f"Error when trying to load tasks: {e}")
        raise

def group_tasks_by_group(tasks):
    """
    Groups the tasks dataframe by team and task_group.
    """
    grouped = tasks.groupby(by=['Responsable', 'Fase']).agg({
        'Fecha Inicio': 'min',
        'Fecha Fin': 'max'
    }).reset_index().sort_values(by=['Fecha Inicio', 'Fase'], ascending=False)
    return grouped

def build_week_ticks(start_date, end_date):
    """
    Identifies the monday dates that will be ticked
    """
    mondays = pd.date_range(start=start_date, end=end_date, freq='W-MON')
    return mondays, [d.strftime('%d') for d in mondays]

def plot_gantt(tasks, output_path=None):
    """
    Plots the Gantt chart and displays it or saves it to output_path.
    """
    if tasks.empty:
        print("No tasks to plot.")
        return
    
    fig, ax = plt.subplots(figsize=(12, 6))
    start_date = tasks['Fecha Inicio'].min()
    end_date = tasks['Fecha Fin'].max()

    tasks = tasks.sort_values(by=['Fecha Inicio', 'Fase'], ascending=False)
    bars = []

    for _, task in tasks.iterrows():
        duration = (task['Fecha Fin'] - task['Fecha Inicio']).days
        label = task.get('Tareas', task['Fase'])
        bar = ax.barh(
                label,
                width=duration,
                height=0.6,
                left=task['Fecha Inicio'],
                color=TEAM_BAR_COLORS.get(task['Responsable'], BAR_COLOR)
        )
        
        for rect in bar:
            rect.annotation_text = (
                f"{task['Fecha Inicio'].strftime('%d/%b/%y')} - {task['Fecha Fin'].strftime('%d/%b/%y')}\n"
                f"Duración: {duration} días\n"
                f"Fase: {task['Fase']}\n"
                f"Responsable: {task['Responsable']}"
            )
            bars.append(rect)
    
    annot = ax.annotate(
        "",
        xy=(0, 0),
        xytext=(10, 10),
        textcoords="offset points",
        bbox=dict(boxstyle="round,pad=0.5", facecolor="white", alpha=0.9, edgecolor="gray"),
        fontsize=10,
        visible=False,
        zorder=100
    )

    def on_hover(event):
        """Maneja el evento de hover sobre las barras"""
        if event.inaxes != ax:
            annot.set_visible(False)
            fig.canvas.draw_idle()
            return
        
        found = False
        for rect in bars:
            if rect.contains(event)[0]:
                # Obtener posición de la barra
                x = rect.get_x() + rect.get_width() / 2
                y = rect.get_y() + rect.get_height() / 2
                
                annot.xy = (x, y)
                annot.set_text(rect.annotation_text)
                annot.set_visible(True)
                found = True
                break
        
        if not found:
            annot.set_visible(False)
        
        fig.canvas.draw_idle()

    fig.canvas.mpl_connect('motion_notify_event', on_hover)

    week_positions, week_labels = build_week_ticks(start_date, end_date)

    ax.set_title(TITLE, fontsize=TITLE_SIZE, color=FONT_COLOR).set_fontweight(TITLE_FONT_WEIGHT)
    ax.set_xlabel(X_LABEL, fontsize=LABEL_SIZE, color=FONT_COLOR)
    ax.set_ylabel(Y_LABEL, fontsize=LABEL_SIZE, color=FONT_COLOR)
    ax.tick_params(axis='both', colors=FONT_COLOR)
    ax.set_xticks(week_positions)
    ax.set_xticklabels(week_labels, fontsize=DAY_FONT_SIZE, color=FONT_COLOR)
    ax.grid(axis='x', linestyle='--', alpha=0.4)

    sec_ax = ax.secondary_xaxis('bottom')
    sec_ax.xaxis.set_major_formatter(mdates.DateFormatter('%b/%y'))
    sec_ax.xaxis.set_major_locator(mdates.MonthLocator())
    sec_ax.tick_params(axis='x', labelsize=MONTH_FONT_SIZE, colors=FONT_COLOR)
    sec_ax.spines['bottom'].set_position(('outward', 20))

    # month line formatting
    for label in sec_ax.get_xticklabels():
        label.set_fontsize(MONTH_FONT_SIZE)
        label.set_weight(MONTH_FONT_WEIGHT)
        label.set_color(FONT_COLOR)
   
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
        sec_ax.spines[spine].set_visible(False)

    # legend
    handles = []
    for team, color in TEAM_BAR_COLORS.items():
        patch = mpatches.Patch(color=color, label=team)
        handles.append(patch)

    ax.legend(handles=handles, loc='best', framealpha=0.8)

    plt.tight_layout()

    if output_path:
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close()
    else:
        plt.show()