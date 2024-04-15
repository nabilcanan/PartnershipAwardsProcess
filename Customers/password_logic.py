# This logic is to check who is using the program, so we added an indicator to see who logs in
# We have this code to add to our main if we so choose but for now we wil leave it out
def log_user_access(username):
    with open("user_access_log.txt", "a") as log_file:
        log_file.write(f"{username} accessed the application.\n")


def ask_password():
    # Create a simple password dialog
    password = simpledialog.askstring("Password", "Enter password:", show='*')
    # Check the password (for example, let's assume the correct password is "admin")
    if password == "admin":
        # If the password is correct, ask for the user's name
        username = simpledialog.askstring("Name", "Enter your name:")
        log_user_access(username)
        # Launch the main application window
        main()
    else:
        # If the password is incorrect, show an error message
        tk.messagebox.showerror("Error", "Incorrect password")