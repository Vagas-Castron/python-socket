import os

# Get the path to the user's Local AppData directory
app_data_path = os.getenv('LOCALAPPDATA')

# Define your app's folder path
app_folder = os.path.join(app_data_path, "ReportApp")

# Create the folder if it doesn't already exist
os.makedirs(app_folder, exist_ok=True)

print(f"App folder created at: {app_folder}")
