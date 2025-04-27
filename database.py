import pandas as pd
import tkinter as tk
from tkinter import messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

def load_file():
    try:
        file_path = 'Data1.xlsx'  # Specify the Excel file name
        df = pd.read_excel(file_path)
        plot_data(df)
    except FileNotFoundError:
        messagebox.showerror("Error", "The file 'Data1.xlsx' was not found.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def plot_data(df):
    # Clear previous figures
    plt.clf()

    # Chart 1: Bar chart of student enrollments in courses
    fig1, ax1 = plt.subplots()
    ax1.bar(df['semester'], df['courses'])
    ax1.set_title("Number of Students Enrolled in Courses")
    ax1.set_xlabel("Semester")
    ax1.set_ylabel("Number of Courses")

    # Chart 2: Horizontal bar chart of student count by profession
    fig2, ax2 = plt.subplots()
    ax2.barh(df['age'], df['semester'])
    ax2.set_title("Number of Students by Profession")
    ax2.set_xlabel("Current Semester ")
    ax2.set_ylabel("Age")

    # Chart 3: Pie chart of student concentration by region
    fig3, ax3 = plt.subplots()
    ax3.pie(df['courses'], labels=df['region'], autopct='%1.1f%%')
    ax3.set_title("Student Concentration by Region")

    # Chart 4: Line chart of courses by semester
    fig4, ax4 = plt.subplots()
    ax4.plot(df['semester'], df['courses'], marker='o')
    ax4.set_title("Courses by Semester")
    ax4.set_xlabel("Semester")
    ax4.set_ylabel("Number of Courses")

    # Chart 5: Area chart of courses by age
    fig5, ax5 = plt.subplots()
    ax5.fill_between(df['age'], df['courses'], alpha=0.5)
    ax5.set_title("Courses by Age")
    ax5.set_xlabel("Age")
    ax5.set_ylabel("Number of Courses")

    # Create a window and add charts
    root = tk.Tk()
    root.title("Dashboard")
    root.state('zoomed')

    side_frame = tk.Frame(root, bg="#FFAF72")
    side_frame.pack(side="left", fill="y")

    label = tk.Label(side_frame, text="Dashboard", bg="#4C2A85", fg="#FFF", font=25)
    label.pack(pady=50, padx=20)

    charts_frame = tk.Frame(root)
    charts_frame.pack()

    upper_frame = tk.Frame(charts_frame)
    upper_frame.pack(fill="both", expand=True)

    # Embed the figures in the Tkinter window
    for fig in [fig1, fig2, fig3]:
        canvas = FigureCanvasTkAgg(fig, upper_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side="left", fill="both", expand=True)

    lower_frame = tk.Frame(charts_frame)
    lower_frame.pack(fill="both", expand=True)

    for fig in [fig4, fig5]:
        canvas = FigureCanvasTkAgg(fig, lower_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side="left", fill="both", expand=True)

# Load the data when the application starts
load_file()

# Start the Tkinter main loop
tk.mainloop()