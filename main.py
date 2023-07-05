import os

import cv2
import json

movie_dir = input("Please insert movie folder")
json_dir = input("Please insert output folder")
assert os.path.exists(json_dir), "output folder does not exist"
movie_number = input("Please insert movie number")
segment_length = int(input("Segment length? (in seconds)"))
video_file = os.path.join(movie_dir, f"{movie_number}.mp4")
assert os.path.exists(video_file), "movie does not exist. Check your path"
start_frame = input("start from 0? if yes press enter, if not insert starting frame")
if not start_frame:
    start_frame = "0"
start_frame_str = start_frame+"+"
json_file = os.path.join(json_dir, f"{movie_number}_{start_frame_str}.json")

frame_labels = {"-1": [], "0": [], "1": [], "2": [], "3": [], "4": [], "5": []}  # Dictionary to store labeled frame intervals


def save_labels():
    with open(json_file, "w") as f:
        json.dump(frame_labels, f)


def label_frame(label, start_frame, end_frame):
    frame_labels[label].append((start_frame, end_frame))

def draw_buttons(frame):
    button_width = frame.shape[1] // 5

    button0 = frame[:, :button_width, :]
    button1 = frame[:, button_width:2 * button_width, :]
    button2 = frame[:, 2 * button_width:3 * button_width, :]
    button3 = frame[:, 3 * button_width:4 * button_width, :]
    button4 = frame[:, 4 * button_width:, :]

    cv2.putText(button0, "0", (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, (0, 0, 255), 2)
    cv2.putText(button1, "1", (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, (0, 0, 255), 2)
    cv2.putText(button2, "2", (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, (0, 0, 255), 2)
    cv2.putText(button3, "3", (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, (0, 0, 255), 2)
    cv2.putText(button4, "Replay", (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, (0, 0, 255), 2)

    frame[:, :button_width, :] = button0
    frame[:, button_width:2 * button_width, :] = button1
    frame[:, 2 * button_width:3 * button_width, :] = button2
    frame[:, 3 * button_width:4 * button_width, :] = button3
    frame[:, 4 * button_width:, :] = button4

    return frame


def play_video(start_frame):
    try:
        cap = cv2.VideoCapture(video_file)
        fps = cap.get(cv2.CAP_PROP_FPS)
        frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        interval_frames = int(fps * segment_length)  # Frames in a 10-second interval

        current_frame = start_frame
        is_playing = True
        interval_counter = 1
        start_frame_interval = current_frame
        replay_start_frame = start_frame + 1
        replay_end_frame = start_frame + 1
        # score_input_active = False

        while True:
            if is_playing:
                ret, frame = cap.read()
                if not ret:
                    break

                frame = draw_buttons(frame)

                cv2.imshow("Video Player", frame)

                current_frame += 1

                if current_frame % interval_frames == 0:
                    # if score_input_active:
                    #current_frame -= interval_frames
                    score = input("Your score? (0-3) Press 'r' to replay: ")
                    if score == 'r':
                        replay_start_frame = start_frame_interval
                        replay_end_frame = current_frame
                        current_frame = replay_start_frame
                        # score_input_active = False
                    elif score == 'q':
                        break
                    else:
                        while score not in frame_labels:
                            score = input("Please enter a valid answer. Your score? (0-3): ")

                        label_frame(score, start_frame_interval, current_frame)
                        # score_input_active = False
                    # else:
                    #     score_input_active = True

                    interval_counter += 1
                    start_frame_interval = current_frame

            key = cv2.waitKey(int(1000 / fps)) & 0xFF

            if key == ord('0'):  # Label frame as 0
                label_frame(0, start_frame_interval, current_frame )
                score_input_active = False

            if key == ord('1'):  # Label frame as 1
                label_frame(1, start_frame_interval, current_frame)
                score_input_active = False

            if key == ord('2'):  # Label frame as 2
                label_frame(2, start_frame_interval, current_frame)
                score_input_active = False

            if key == ord('3'):  # Label frame as 3
                label_frame(3, start_frame_interval, current_frame)
                score_input_active = False

            if key == ord('r'):  # Replay the last section
                replay_start_frame = start_frame_interval
                replay_end_frame = current_frame
                current_frame = replay_start_frame
                score_input_active = False

            if replay_start_frame <= current_frame <= replay_end_frame:  # Check if in replay mode
                cap.set(cv2.CAP_PROP_POS_FRAMES, replay_start_frame)
                replay_start_frame += 1

            if key == ord('p'):  # Pause or resume playback
                is_playing = not is_playing

            if key == ord('q'):  # Quit the video player
                break

            if current_frame >= frame_count:
                break

        cap.release()
        cv2.destroyAllWindows()

    except KeyboardInterrupt as kbi:
        print(kbi)

    except Exception as ex:
        print(ex)

    save_labels()

play_video(int(start_frame))  # Start the video from frame 0