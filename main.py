import gui
from converter import *
import argparse

if __name__ == '__main__':
    # ws = gui.generate_window()
    # ws.mainloop()

    parser = argparse.ArgumentParser("simple_example")
    parser.add_argument("-f", "--filename", required=True, help="The path to the openlp service (*.osz) file.", type=str)
    parser.add_argument("-b","--background", default=DEFAULT_THEME, choices=["light", "dark"], help="Whether the background images should be light or dark", type=str)
    args = parser.parse_args()

    job = Job(osz_file_path=args.filename, theme=args.background)
    # job = Job(osz_file_path="Service 2021-10-10.osz", theme="light")
    job.get_songs_from_osz()
    job.generate_ppt()
    job.save_file()

    print("...done")
