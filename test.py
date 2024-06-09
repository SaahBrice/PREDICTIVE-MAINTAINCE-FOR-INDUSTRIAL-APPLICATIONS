import pandas as pd
import plotly.graph_objs as go
from plotly.subplots import make_subplots

# Load the data from the provided Excel file
file_path = 'data.xlsx'
df = pd.read_excel(file_path)

# Strip leading and trailing whitespace from column names
df.columns = df.columns.str.strip()

# Convert HP column to numeric
df['HP'] = pd.to_numeric(df['HP'], errors='coerce')

# Adjust the column names in the plotting code
def generate_plot(df, time_step):
    fig = make_subplots(rows=2, cols=3, subplot_titles=("Vibration Level", "Vibration Intensity", "Sound Level", "Sound Value", "HP", "Life Span"))

    time_filtered_df = df[df['Time'] <= time_step]

    fig.add_trace(go.Scatter(x=time_filtered_df['Time'], y=time_filtered_df['Vibration Level'], mode='lines+markers', name='Vibration Level'), row=1, col=1)
    fig.add_trace(go.Scatter(x=time_filtered_df['Time'], y=time_filtered_df['Vibration Intensity'], mode='lines+markers', name='Vibration Intensity'), row=1, col=2)
    fig.add_trace(go.Scatter(x=time_filtered_df['Time'], y=time_filtered_df['Sound Level'], mode='lines+markers', name='Sound Level'), row=1, col=3)
    fig.add_trace(go.Scatter(x=time_filtered_df['Time'], y=time_filtered_df['Sound Value'], mode='lines+markers', name='Sound Value'), row=2, col=1)
    fig.add_trace(go.Scatter(x=time_filtered_df['Time'], y=time_filtered_df['HP'], mode='lines+markers', name='HP'), row=2, col=2)
    fig.add_trace(go.Scatter(x=time_filtered_df['Time'], y=time_filtered_df['Life Span'], mode='lines+markers', name='Life Span'), row=2, col=3)

    fig.update_layout(title_text=f"ICT FOR INDUSTRIAL AUTOMATION FINAL PROJECT Data Visualization")
    return fig

# Generate initial of all plots
fig = generate_plot(df, 1)

# Add shapes to highlight HP regions to the initial plot layout
fig.update_layout(
    shapes=[
        # Green region for HP > 50
        dict(type='rect', xref='x5', yref='y5', x0=df['Time'].min(), y0=50, x1=df['Time'].max(), y1=100, fillcolor='green', opacity=0.2, layer='below', line_width=0),
        # Yellow region for 20 < HP <= 50
        dict(type='rect', xref='x5', yref='y5', x0=df['Time'].min(), y0=20, x1=df['Time'].max(), y1=50, fillcolor='yellow', opacity=0.2, layer='below', line_width=0),
        # Red region for HP <= 20
        dict(type='rect', xref='x5', yref='y5', x0=df['Time'].min(), y0=0, x1=df['Time'].max(), y1=20, fillcolor='red', opacity=0.2, layer='below', line_width=0),
    ]
)

# Generate frames for the animation
frames = []
for t in df['Time']:
    time_filtered_df = df[df['Time'] <= t]
    frame_fig = generate_plot(df, t)
    frame_shapes = [
        dict(type='rect', xref='x5', yref='y5', x0=time_filtered_df['Time'].min(), y0=50, x1=time_filtered_df['Time'].max(), y1=300, fillcolor='green', opacity=0.2, layer='below', line_width=0),
        dict(type='rect', xref='x5', yref='y5', x0=time_filtered_df['Time'].min(), y0=20, x1=time_filtered_df['Time'].max(), y1=50, fillcolor='yellow', opacity=0.2, layer='below', line_width=0),
        dict(type='rect', xref='x5', yref='y5', x0=time_filtered_df['Time'].min(), y0=0, x1=time_filtered_df['Time'].max(), y1=20, fillcolor='red', opacity=0.2, layer='below', line_width=0),
    ]
    frame_fig.update_layout(shapes=frame_shapes)
    frames.append(go.Frame(data=frame_fig.data, layout=frame_fig.layout, name=str(t)))

# Assign frames to the initial plot
fig.frames = frames

# Add play button
fig.update_layout(
    updatemenus=[dict(
        type="buttons",
        showactive=False,
        buttons=[dict(label="Play",
                      method="animate",
                      args=[None, dict(frame=dict(duration=500, redraw=True), fromcurrent=True)])]
    )]
)

# Save the plot as an HTML file
fig.write_html("interactive_plot.html")

# Show the plot in an interactive window (if running in a Jupyter environment)
fig.show()
