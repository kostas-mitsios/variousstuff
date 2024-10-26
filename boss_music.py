import os
import pygame
import time

def play_mp3(file_name, loop_count):
    file_path = os.path.join(os.getcwd(), file_name)
    
    pygame.mixer.init(channels=0)  # Set channels to zero to suppress the pygame message
    pygame.mixer.music.load(file_path)

    for _ in range(loop_count):
        pygame.mixer.music.play()
        # Wait for the MP3 to finish playing
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(2)  # You can adjust the tick value if needed.

    pygame.mixer.quit()

if __name__ == "__main__":
    mp3_file_name = "boss_music.mp3"  # Replace with the name of your MP3 file
    loop_count = 3

    play_mp3(mp3_file_name, loop_count)
