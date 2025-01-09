import os
import pickle
import tkinter as tk
from tkinter import messagebox
import webbrowser
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_script_dir():
    """Get the directory of the current script."""
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()


def load_settings():
    """Load settings from the 'settings.pkl' file."""
    settings_file = os.path.join(get_script_dir(), 'settings.pkl')
    if os.path.exists(settings_file):
        with open(settings_file, 'rb') as f:
            settings = pickle.load(f)
            if 'user_agent' not in settings:
                settings['user_agent'] = ''
            if 'urls' not in settings:
                settings['urls'] = []
            if 'proxies' not in settings:
                settings['proxies'] = []
            return settings
    else:
        return {
            'user_agent': '',
            'urls': [],
            'proxies': []  # Format: [{'ip': '...', 'login': '...', 'password': '...'}]
        }


def save_settings(settings):
    """Save settings to the 'settings.pkl' file."""
    settings_file = os.path.join(get_script_dir(), 'settings.pkl')
    with open(settings_file, 'wb') as f:
        pickle.dump(settings, f)


def check_google_indexation(settings):
    """Check if URLs are indexed in Google."""
    headers = {'User-Agent': settings['user_agent']}
    urls = settings['urls']
    proxies = settings['proxies']
    use_proxies = bool(proxies)  # Check if proxies are provided
    current_proxy_index = 0  # Index of the current proxy, if available

    results = {}
    for url in urls:
        while True:
            try:
                # Configure proxy with login and password
                if use_proxies:
                    proxy_info = proxies[current_proxy_index]
                    proxy = {
                        'http': f"http://{proxy_info['login']}:{proxy_info['password']}@{proxy_info['ip']}",
                        'https': f"http://{proxy_info['login']}:{proxy_info['password']}@{proxy_info['ip']}"
                    }
                else:
                    proxy = None

                search_url = f"https://www.google.com/search?q=site:{url}"
                response = requests.get(search_url, headers=headers, proxies=proxy, timeout=10)
                response.raise_for_status()

                soup = BeautifulSoup(response.text, 'html.parser')
                search_results = soup.find_all('div', class_='tF2Cxc')

                if search_results:
                    results[url] = "Indexed"
                else:
                    results[url] = "Not Indexed"
                break  # Exit loop on successful check
            except requests.exceptions.ProxyError:
                if use_proxies:
                    current_proxy_index += 1
                    if current_proxy_index >= len(proxies):
                        results[url] = "Error: All proxies failed"
                        break
                else:
                    results[url] = "Error: No proxy available"
                    break
            except requests.exceptions.Timeout:
                results[url] = "Error: Request timeout"
                break
            except requests.exceptions.RequestException as e:
                results[url] = f"Error: {str(e)}"
                break

    # Save the results to an Excel file
    save_results_to_excel(results)


def save_results_to_excel(results):
    """Save the results to an Excel file."""
    script_dir = get_script_dir()
    output_file = os.path.join(script_dir, "indexation_results.xlsx")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Indexation Results"
    sheet.append(["URL", "Status"])

    for url, status in results.items():
        sheet.append([url, status])

    workbook.save(output_file)
    messagebox.showinfo("Information", f"Results saved to {output_file}")


def create_gui(settings):
    """Create the GUI for the script."""
    window = tk.Tk()
    window.title("Google Index Checker")

    # User-Agent
    tk.Label(window, text="User Agent:").grid(row=0, column=0, sticky="w")
    user_agent_entry = tk.Entry(window, width=50)
    user_agent_entry.insert(0, settings['user_agent'])
    user_agent_entry.grid(row=0, column=1)

    # Add a blue clickable link for User-Agent help
    link_label = tk.Label(window, text="Find your user agent at https://www.whatsmyua.info", fg="blue", cursor="hand2")
    link_label.grid(row=1, column=1, sticky="w")
    link_label.bind("<Button-1>", lambda e: webbrowser.open("https://www.whatsmyua.info"))

    # URLs
    tk.Label(window, text="URLs (one per line, up to 1000):").grid(row=2, column=0, sticky="nw")
    urls_text = tk.Text(window, height=15, width=50)
    urls_text.insert('1.0', '\n'.join(settings['urls']))
    urls_text.grid(row=2, column=1, padx=5, pady=5)

    # Proxies
    tk.Label(window, text="Proxies (IP, Login, Password):").grid(row=3, column=0, sticky="nw")
    proxies_frame = tk.Frame(window)
    proxies_frame.grid(row=3, column=1, padx=5, pady=5)

    # Column headers for proxies
    tk.Label(proxies_frame, text="IP Address", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=2)
    tk.Label(proxies_frame, text="Login", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=2)
    tk.Label(proxies_frame, text="Password", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=2)

    proxy_ip_entries = []
    proxy_login_entries = []
    proxy_password_entries = []

    for i in range(5):  # Allow up to 5 proxies
        ip_entry = tk.Entry(proxies_frame, width=20)
        ip_entry.grid(row=i + 1, column=0, padx=5, pady=2)
        proxy_ip_entries.append(ip_entry)

        login_entry = tk.Entry(proxies_frame, width=15)
        login_entry.grid(row=i + 1, column=1, padx=5, pady=2)
        proxy_login_entries.append(login_entry)

        password_entry = tk.Entry(proxies_frame, width=15, show="*")
        password_entry.grid(row=i + 1, column=2, padx=5, pady=2)
        proxy_password_entries.append(password_entry)

    # Load proxies into the interface
    for i, proxy in enumerate(settings['proxies']):
        if i < 5:
            proxy_ip_entries[i].insert(0, proxy.get('ip', ''))
            proxy_login_entries[i].insert(0, proxy.get('login', ''))
            proxy_password_entries[i].insert(0, proxy.get('password', ''))

    def on_save():
        """Save settings and start the indexation check."""
        settings['user_agent'] = user_agent_entry.get()
        settings['urls'] = urls_text.get('1.0', tk.END).strip().split('\n')

        proxies = []
        for i in range(5):
            ip = proxy_ip_entries[i].get().strip()
            login = proxy_login_entries[i].get().strip()
            password = proxy_password_entries[i].get().strip()
            if ip:
                proxies.append({'ip': ip, 'login': login, 'password': password})

        settings['proxies'] = proxies
        save_settings(settings)
        check_google_indexation(settings)

    tk.Button(window, text="Save Settings and Run Search", command=on_save).grid(row=4, column=1, pady=10)

    window.mainloop()


if __name__ == "__main__":
    settings = load_settings()
    create_gui(settings)
