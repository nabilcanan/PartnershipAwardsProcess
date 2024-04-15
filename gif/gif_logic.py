from PIL import Image, ImageTk
import os
import platform
import subprocess


def open_powerpoint():
    filepath = "P:\\Partnership_Python_Projects\\TTI Shared Data Program\\TTI_SHARED_DATA_INSTRUCTIONS.pptx"

    if platform.system() == 'Windows':
        os.startfile(filepath)
    else:
        subprocess.call(['open', filepath])


# from CREATION_DND import creation_logic_dnd
# Not going to import now, working on fixed version of this module


def update_background(canvas, frames, current=0, width=1000, height=650):
    frame = frames[current]
    # Use a dictionary to store a reference to the image, preventing garbage collection.
    if not hasattr(canvas, 'image_references'):
        canvas.image_references = {}  # Initialize if it doesn't exist
    # Store the reference with the current index then move them accordingly
    canvas.image_references[current] = frame
    canvas.itemconfig("background", image=frame)
    current = (current + 1) % len(frames)
    canvas.after(100, update_background, canvas, frames, current, width, height)


def load_gif_frames(filename, width, height):
    frames = []
    with Image.open(filename) as im:
        while True:
            try:
                im.seek(im.tell())
                # Use Image.LANCZOS or another available filter directly
                frame_image = im.copy().resize((width, height), Image.LANCZOS)
                frames.append(ImageTk.PhotoImage(frame_image))
                im.seek(im.tell() + 1)
            except EOFError:
                break  # Exit the loop when there are no more frames. There are continuous frames in this gif
    return frames


def update_animation(canvas, frames, position, step, direction, current_frame=0):
    if position < 0 or position > canvas.winfo_width() - frames[0].width():
        direction *= -1  # Change direction
    position += step * direction
    frame = frames[current_frame % len(frames)]
    canvas.itemconfig('animation', image=frame)
    canvas.coords('animation', position, canvas.winfo_height() - frames[0].height())
    current_frame += 1
    canvas.after(100, update_animation, canvas, frames, position, step, direction, current_frame)


def load_gif(filename):
    frames = []
    with Image.open(filename) as im:
        while True:
            try:
                im.seek(im.tell())
                frames.append(ImageTk.PhotoImage(im.copy()))
                im.seek(im.tell() + 1)
            except EOFError:
                break
    return frames
