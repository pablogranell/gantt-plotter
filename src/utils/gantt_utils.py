import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.patches as mpatches
from matplotlib.widgets import Button

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
rcParams['toolbar'] = 'None'

# Columnas esperadas del Excel
COLUMN_NAMES = ['Fase', 'Tareas', 'Responsable', 'Duración', 'Fecha Inicio', 'Fecha Fin']

hover_enabled = True


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
            skiprows=skiprows
        )

        tasks.columns = COLUMN_NAMES
        tasks['Fecha Inicio'] = pd.to_datetime(tasks['Fecha Inicio'], format=DATE_FORMAT)
        tasks['Fecha Fin'] = pd.to_datetime(tasks['Fecha Fin'], format=DATE_FORMAT)
        tasks.set_index(pd.DatetimeIndex(tasks['Fecha Inicio'].values), inplace=True)
        
        return tasks
    
    except Exception as e:
        print(f"Error al cargar las tareas: {e}")
        raise

def group_tasks_by_group(tasks):
    grouped = tasks.groupby(by=['Responsable', 'Fase']).agg({
        'Fecha Inicio': 'min',
        'Fecha Fin': 'max'
    }).reset_index().sort_values(by=['Fecha Inicio', 'Fase'], ascending=False)
    return grouped

def build_week_ticks(start_date, end_date):
    """Identifica los lunes que se marcarán en el eje X."""
    mondays = pd.date_range(start=start_date, end=end_date, freq='W-MON')
    labels = [d.strftime('%d') for d in mondays]
    return mondays, labels

def _create_annotation(ax):
    return ax.annotate(
        "",
        xy=(0, 0),
        xytext=(10, 10),
        textcoords="offset points",
        bbox=dict(
            boxstyle="round,pad=0.5",
            facecolor="white",
            alpha=0.9,
            edgecolor="gray"
        ),
        fontsize=10,
        visible=False,
        zorder=100
    )


def _format_bar_annotation(task, duration):
    return (
        f"{task['Fecha Inicio'].strftime('%d/%b/%y')} - "
        f"{task['Fecha Fin'].strftime('%d/%b/%y')}\n"
        f"{duration} días\n"
        f"{task['Fase']}\n"
        f"{task['Responsable']}"
    )


def _setup_hover_handler(fig, ax, bars, annot):    
    def on_hover(event):
        """Maneja el evento de hover sobre las barras"""
        if not hover_enabled:
            annot.set_visible(False)
            fig.canvas.draw_idle()
            return
            
        if event.inaxes != ax:
            annot.set_visible(False)
            fig.canvas.draw_idle()
            return
        
        for rect in bars:
            if rect.contains(event)[0]:
                # Obtener posición de la barra
                x = rect.get_x() + rect.get_width() / 2
                y = rect.get_y() + rect.get_height() / 2
                
                annot.xy = (x, y)
                annot.set_text(rect.annotation_text)
                annot.set_visible(True)
                fig.canvas.draw_idle()
                return

        annot.set_visible(False)
        fig.canvas.draw_idle()

    fig.canvas.mpl_connect('motion_notify_event', on_hover)


def _configure_axes(ax, sec_ax, week_positions, week_labels):
    # Título y etiquetas
    ax.set_title(TITLE, fontsize=TITLE_SIZE, color=FONT_COLOR).set_fontweight(TITLE_FONT_WEIGHT)
    ax.set_xlabel(X_LABEL, fontsize=LABEL_SIZE, color=FONT_COLOR)
    ax.set_ylabel(Y_LABEL, fontsize=LABEL_SIZE, color=FONT_COLOR)
    # Eje X primario (semanas)
    ax.tick_params(axis='both', colors=FONT_COLOR)
    ax.set_xticks(week_positions)
    ax.set_xticklabels(week_labels, fontsize=DAY_FONT_SIZE, color=FONT_COLOR)
    ax.grid(axis='x', linestyle='--', alpha=0.4)

    # Eje X secundario (meses)
    sec_ax.xaxis.set_major_formatter(mdates.DateFormatter('%b/%y'))
    sec_ax.xaxis.set_major_locator(mdates.MonthLocator())
    sec_ax.tick_params(axis='x', labelsize=MONTH_FONT_SIZE, colors=FONT_COLOR)
    sec_ax.spines['bottom'].set_position(('outward', 20))
    for label in sec_ax.get_xticklabels():
        label.set_fontsize(MONTH_FONT_SIZE)
        label.set_weight(MONTH_FONT_WEIGHT)
        label.set_color(FONT_COLOR)
    # Ocultar bordes innecesarios
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
        sec_ax.spines[spine].set_visible(False)


def _create_legend(ax):
    handles = [
        mpatches.Patch(color=color, label=team)
        for team, color in TEAM_BAR_COLORS.items()
    ]
    ax.legend(handles=handles, loc='best', framealpha=0.8)


def _create_floating_buttons(fig):
    global hover_enabled
    buttons = []
    button_axes = []
    
    # Botón Guardar
    ax_save = fig.add_axes([0.01, 0.94, 0.07, 0.05])
    btn_save = Button(ax_save, 'Guardar', color='white', hovercolor='lightgray')
    btn_save.label.set_fontsize(14)
    button_axes.append(ax_save)
    initial_color = 'palegreen' if hover_enabled else 'white'
    ax_hover = fig.add_axes([0.085, 0.94, 0.12, 0.05])
    btn_hover = Button(ax_hover, 'Menú flotante', color=initial_color)
    btn_hover.label.set_fontsize(14)
    button_axes.append(ax_hover)
    
    def save_click(event):
        filepath = 'gantt.png'
        for ax_btn in button_axes:
            ax_btn.set_visible(False)
        fig.canvas.draw()
        
        fig.savefig(filepath, dpi=300, bbox_inches='tight')
        print(f"Guardado en {filepath}")
        
        # Volver a mostrar botones
        for ax_btn in button_axes:
            ax_btn.set_visible(True)
        fig.canvas.draw_idle()
    
    btn_save.on_clicked(save_click)
    buttons.append(btn_save)
    
    def hover_click(event):
        global hover_enabled
        hover_enabled = not hover_enabled
        new_color = 'palegreen' if hover_enabled else 'white'
        ax_hover.set_facecolor(new_color)
        btn_hover.color = new_color
        fig.canvas.draw_idle()
    
    btn_hover.on_clicked(hover_click)
    buttons.append(btn_hover)
    
    return buttons


def plot_gantt(tasks, output_path=None):
    if tasks.empty:
        print("No hay tareas para dibujar.")
        return

    fig, ax = plt.subplots(figsize=(12, 6))
    ax.format_coord = lambda x, y: ''
    
    start_date = tasks['Fecha Inicio'].min()
    end_date = tasks['Fecha Fin'].max()
    tasks = tasks.sort_values(by=['Fecha Inicio', 'Fase'], ascending=False)

    bars = []
    for _, task in tasks.iterrows():
        duration = (task['Fecha Fin'] - task['Fecha Inicio']).days
        label = task.get('Tareas', task['Fase'])
        color = TEAM_BAR_COLORS.get(task['Responsable'], BAR_COLOR)

        bar = ax.barh(label, width=duration, height=0.6, left=task['Fecha Inicio'], color=color)

        for rect in bar:
            rect.annotation_text = _format_bar_annotation(task, duration)
            bars.append(rect)

    annot = _create_annotation(ax)
    _setup_hover_handler(fig, ax, bars, annot)

    week_positions, week_labels = build_week_ticks(start_date, end_date)
    sec_ax = ax.secondary_xaxis('bottom')
    
    _configure_axes(ax, sec_ax, week_positions, week_labels)
    _create_legend(ax)
    
    if not output_path:
        buttons = _create_floating_buttons(fig)
        fig._buttons = buttons

    plt.tight_layout()

    if output_path:
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close()
    else:
        plt.show()