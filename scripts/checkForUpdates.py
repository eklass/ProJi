import subprocess

def check_for_updates():
    # Prüfen, ob es Updates gibt
    subprocess.run(["git", "fetch"], check=True)
    result = subprocess.run(["git", "status", "-uno"], capture_output=True, text=True)
    if "behind" in result.stdout:
        show_applescript_popup()

def show_applescript_popup():
    script = """
    display dialog "Ein Update ist verfügbar!\nBitte aktualisieren Sie Ihr Repository." buttons {"OK"} default button "OK"
    """
    subprocess.run(["osascript", "-e", script], check=True)


if __name__ == "__main__":
    check_for_updates()
