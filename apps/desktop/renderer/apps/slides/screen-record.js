// Insert → Screen Recording and Record → Screen Recording are the suite's
// Recorder now (recording/open-recorder.js, its 'slide' mode): the
// window, screen or region picked is recorded until Stop, and the recording
// comes back as a video for the slide. What stays here is the container a
// slide's recording is made in, which the cameo and the video export use too,
// and the opening of a recorder in it.

export { recordingType, openRecorder } from '../../recording/session.js';
