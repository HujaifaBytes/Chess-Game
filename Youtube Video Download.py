from yt_dlp import YoutubeDL
def download_video(url, save_path):
    ydl_opts = {
        'outtmpl': f'{save_path}/%(title)s.%(ext)s',
        'format': 'bestvideo+bestaudio/best',
        'merge_output_format': 'mp4',
    }
    with YoutubeDL(ydl_opts) as ydl:
        ydl.download([url])
        print("Video downloaded successfully!")
        
url = input("Enter the YouTube video URL: ")
save_path = r"E:\python Auto"

download_video(url, save_path)