import pandas as pd
import plotly.graph_objs as go
from plotly.subplots import make_subplots

# Load the data from the provided Excel file
file_path = 'DataExcel.xlsx'
df = pd.read_excel(file_path)

# Strip leading and trailing whitespace from column names
df.columns = df.columns.str.strip()

# Print the column names to verify them
print(df.columns)

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

    fig.update_layout(height=600, width=900, title_text=f"ICT FOR INDUSTRIAL AUTOMATION FINAL PROJECT Data Visualization")
    return fig

# Generate frames for the animation
frames = [go.Frame(data=generate_plot(df, t).data, name=str(t)) for t in df['Time']]

# Initial plot
fig = generate_plot(df, 1)
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
