import tkinter as tk


#  Flextronics
def flextronics_logic():
    # Create a new Toplevel window for Benchmark
    benchmark_window = tk.Toplevel()
    benchmark_window.title("Flextronics Module")
    benchmark_window.geometry("600x400")

    # Add content specific to the Benchmark module
    label = tk.Label(benchmark_window, text="This is the Flextronics module", font=("Verdana", 16))
    label.pack(pady=20)
