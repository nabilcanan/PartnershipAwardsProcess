from Customers.Benchmark import benchmark_logic
from Customers.Creation import creation_logic
from Customers.Flextronics import flextronics_logic
from Customers.Jabil import jabil_logic
from Customers.Kimball import kimball_logic
from Customers.Neotech import neotech_logic
from Customers.Plexus import plexus_logic
from Customers.Sanmina import sanmina_logic
from Customers.SMTC import smtc_logic
import tkinter as tk
from tkinter import Canvas  # , simpledialog <--- add this line if you want to add a password and uncomment
from gif.gif_logic import *


def button_click(customer_name):
    # Functionality to be executed when a label/button is clicked
    open_customer_module(customer_name)


def open_customer_module(customer_name):
    # Define a dictionary mapping customer names to their corresponding functions
    customer_logic = {
        "Benchmark": benchmark_logic,
        "Creation": creation_logic,
        "Flextronics": flextronics_logic,
        "Jabil": jabil_logic,
        "Kimball": kimball_logic,
        "Neotech": neotech_logic,
        "Plexus": plexus_logic,
        "Sanmina": sanmina_logic,
        "SMTC": smtc_logic,
        # "Kauffman Engineering": kauffman_logic not yet used for any function or logic
    }

    # Get the corresponding function based on customer name and call it
    logic_function = customer_logic.get(customer_name)
    if logic_function:
        logic_function()
    else:
        # Handle the case where the customer name is not found
        print(f"Customer '{customer_name}' not recognized.")


def main():
    root = tk.Tk()
    root.title("Partnership TTI Program")
    root.geometry("1000x550")

    # Add a button for opening the PowerPoint instructions
    instructions_button = tk.Button(root, text="Open Instructions", command=open_powerpoint)
    instructions_button.pack()

    # Modern color scheme
    bg_color = '#2c3e50'  # Dark blue
    button_color = '#34495e'  # Greyish blue
    button_hover_color = '#4a6fa5'  # Lighter blue
    text_color = '#ecf0f1'  # Light grey

    root.configure(bg=bg_color)

    # Create canvas for background and buttons
    canvas = Canvas(root, width=1000, height=650, bg=bg_color, highlightthickness=0)
    canvas.pack(fill="both", expand=True)

    # Title above the buttons
    canvas.create_text(500, 50, text="TTI Shared Data", font=('Roboto', 32), fill=text_color)

    # Customers list
    customers = ["Benchmark", "Creation", "Jabil",
                 "Kimball", "Neotech", "Plexus", "Sanmina"]

    # Functionality for button hover effect
    def on_enter(tag):
        canvas.itemconfig(tag, fill=button_hover_color)

    def on_leave(tag):
        canvas.itemconfig(tag, fill=button_color)

    # Create buttons using text objects on the canvas, for each customer
    button_y = 100
    button_height = 40
    button_width = 300
    for customer in customers:
        tag_rect = f'rect_{customer}'
        # Print out each of the customer tags in the terminal and display which one we are opening
        tag_text = f'text_{customer}'

        # Create rectangle as button background
        canvas.create_rectangle(500 - button_width // 2, button_y - button_height // 2,
                                500 + button_width // 2, button_y + button_height // 2,
                                fill=button_color, outline=text_color, tags=(customer, tag_rect))

        # Create text on top of the rectangle
        canvas.create_text(500, button_y, text=customer, font=('Arial', 16), fill=text_color,
                           tags=(customer, tag_text))

        # Bind the click event to the rectangle and text
        canvas.tag_bind(customer, '<Button-1>', lambda event, name=customer: button_click(name))

        # Bind the hover effects
        canvas.tag_bind(customer, '<Enter>', lambda e, tag=tag_rect: on_enter(tag))
        canvas.tag_bind(customer, '<Leave>', lambda e, tag=tag_rect: on_leave(tag))

        button_y += 60  # Increment y position for the next button, this adds spacing

        # Load GIF frames and create the animation
        gif_frames = load_gif('gif/component.gif')
        canvas.create_image(0, root.winfo_height() - gif_frames[0].height(), anchor='nw', tags='animation',
                            image=gif_frames[0])
        update_animation(canvas, gif_frames, 0, 10, 1)  # Start animation

    root.mainloop()


if __name__ == "__main__":
    main()
