import sys
import pandas as pd
import os
import yt_dlp as youtube_dl

#  HANDLING ARGUMENTS
#  Get the folder name and Excel sheet name from command-line arguments
if len(sys.argv) == 4:
    output_dir = sys.argv[3]
    excel_sheet_name = sys.argv[2]
    folder_name = sys.argv[1]
elif len(sys.argv) == 3:
    output_dir = None
    folder_name = sys.argv[1]
    excel_sheet_name = sys.argv[2]
else:
    print('---------------- INVALID USAGE ------------------')
    print('3 ARGUMENT Usage: python3 makecrate.py [crate_name] [excel_sheet_name].xlsx [output_directory ]')
    print("2 ARGUMENT Usage: python3 makecrate.py [crate_name] [excel_sheet_name].xlsx ")
    sys.exit()


# Read the Excel sheet containing song titles and authors
df = pd.read_excel(excel_sheet_name)


# Create path to folder in output directory
folder_path = os.path.join(output_dir , folder_name) if output_dir else os.path.join(folder_name)
# Check if the folder already exists
if os.path.exists(folder_path):
    print('Folder already exists, songs will be downloaded to:', folder_path)
else:
    print('Folder does not exist, creating new folder:', folder_path)
    os.makedirs(folder_path)

# Loop through each row in the Excel sheet
for index, row in df.iterrows():
    # Construct the YouTube search query using song title and author
    query = row['Song Title'] + ' ' + row['Author'] + ' audio'
    print('Searching for:', query)
    # Set the output paths to the folder where you want to save the songs to verify no duplicates
    output_path = os.path.join(folder_path, row['Song Title'] + ".mp3")
    output_filename = os.path.join(folder_path, row['Song Title'])

    print(type(row['Rename']))
    if type(row['Rename']) == str: #if rename value assigned in sheet
        print("renaming to", row['Rename'])
        rename = str(row['Rename'])
        output_path = os.path.join(folder_path, rename + ".mp3")
        output_filename = os.path.join(folder_path, rename)


    if os.path.exists(output_path):
         print('Song already exists, skipping download:', row['Song Title'] + ".mp3")
    else:
        # Search for the song on YouTube
        ydl_opts = {'format': 'bestaudio/best', 'outtmpl': output_filename, 'postprocessors': [{'key': 'FFmpegExtractAudio','preferredcodec': 'mp3','preferredquality': '192',}],}
        with youtube_dl.YoutubeDL(ydl_opts) as ydl:
            ydl.download(['ytsearch:' + query])

        print('Song Successfully Downloaded:', row['Song Title'])

print(f'Successfully Downloaded all songs to {folder_name},! Time to MIX...')
