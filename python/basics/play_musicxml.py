import os
import tempfile
import time

import partitura as pt
import pygame


def play_musicxml(musicxml_file_path):
    """
    Parses a MusicXML file and plays it using pygame.

    Args:
        musicxml_file_path (str): Path to the MusicXML file.
    """
    try:
        # 1. Parse the MusicXML file into partitura's internal structure
        # This will automatically handle the conversion of notation to
        # performable events (like unrolling repeats)
        score = pt.load_score(musicxml_file_path)
        print(f"Successfully parsed score: {score.title}")

        # 2. Extract a Part object to convert to a performance (MIDI-like data)
        # We can take the first part as an example
        if not score.parts:
            print("Score has no parts to play.")
            return

        part = score.parts[0]  # You might want to iterate or handle all parts

        # 3. Convert the symbolic representation to a sequence of events (like MIDI messages)
        # This creates a "performance" object
        performance = pt.musicanalysis.performance_from_score(part)

        # 4. Save the performance as a temporary MIDI file
        with tempfile.NamedTemporaryFile(suffix=".mid", delete=False) as tmpfile:
            temp_midi_path = tmpfile.name

        pt.save_performance_midi(performance, temp_midi_path)
        print(f"Temporary MIDI file saved to: {temp_midi_path}")

        # 5. Use pygame to play the temporary MIDI file
        pygame.init()
        pygame.mixer.init()

        try:
            pygame.mixer.music.load(temp_midi_path)
            print("Starting playback...")
            pygame.mixer.music.play()

            while pygame.mixer.music.get_busy():
                time.sleep(1)  # Wait while music plays

        except pygame.error as e:
            print(
                "Could not play MIDI file: "
                f"{e}. Ensure you have necessary MIDI software/drivers "
                "(e.g., timidity or a soundfont) installed on your system for pygame to use."
            )

        finally:
            # 6. Clean up the temporary file
            pygame.mixer.music.stop()
            pygame.mixer.quit()
            pygame.quit()
            os.remove(temp_midi_path)
            print(f"Cleaned up {temp_midi_path}")

    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    # Example Usage:
    # Replace 'your_score.musicxml' with the path to your actual MusicXML file
    # play_musicxml("your_score.musicxml")
    pass
