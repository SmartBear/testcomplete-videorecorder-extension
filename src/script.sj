// Log messages
var logMessages = {
  recorderIsNotInstalled: {
    message: "Unable to start video recording. The VLC recorder is not installed. See Additional Info for details.",
    messageEx: "<p>The video recorder uses the recording functionality of the VLC framework.<br/>Download the recorder from the following website and install it:<br/><a href='%s' target='_blank'>%s</a></p>"
  },
  startOk: {
    message: "Video recording has started. The recorded video will be saved to the <your-project>\Logs folder. You will find the file name on the Additional Info tab.",
    messageEx: "Video quality: %s\r\n\r\nThe name of the video file to be created:\r\n%s"
  },
  startFailAlreadyStarted: {
    message: "Unable to start video recording. The VLC recorder is already running. See Additional Info for details.",
    messageEx: "You need to stop the running instance of the VLC recorder before starting a new video recording session.\r\nIf you see the " + "%s" + ".exe process in the system, terminate it."
  },
  startNoRecorderProcess: {
    message: "Unable to start the video recorder. See Additional Info for details.",
    messageEx: "<p>Your test failed to start the video recorder. Something is wrong in the system.<p>" +
      "<p>To get more information:</p>" +
      "<ul>" +
      "<li>Run the VLC recorder in the system. Use the following command line for this:<br/>%s</li>" +
      "<li>Explore the messages in the command-line window.</li>" +
      "</ul>"
  },
  stopOk: {
    message: "Video recording is over. You can find the video file name on the Additional Info tab.",
    messageEx: "Video file name:\r\n%s"
  },
  stopFailNoRecorderProcess: {
    message: "Unable to find the video recorder. See Additional Info for details.",
    messageEx: "Unable to find a running instance of the VLC recorder. Please make sure you have started recording in your test."
  },
  stopFailRecorderNotStarted: {
    message: "Video recording has not been started in your test. See Additional Info for details.",
    messageEx: "Most likely, the running instance of the VLC recorder was not started from your test.\r\nIf you see the %s.exe process running in the system, terminate it."
  },
  recorderUnexpectedError: {
    message: "The video file has not been created. See Additional Info for details.",
    messageEx: "<p>Your test failed to start the video recorder. Something is wrong in the system.<p>" +
      "<p>To get more information:</p>" +
      "<ul>" +
      "<li>Run the VLC recorder in the system. Use the following command line for this:<br/>%s</li>" +
      "<li>Explore the messages in the command-line window.</li>" +
      "</ul>"
  },
  processWasTerminated: {
    message: "The VLC recorder process has been terminated on timeout. See Additional Info for details.",
    messageEx: "The VLC recorder was terminated on timeout. Most likely, it had been saving the video for a long time. Please try to record a shorter video."
  }
};

// Other messages
var messages = {
  encodingInProgress: "Encoding the video file..."
};

// Creates an object that provides info on the recording engine
function RecorderInfo() {
  this.getHomepage = function () {
    return "https://www.videolan.org/";
  };

  this.getProcessName = function () {
    return "vlc";
  };

  this.doesProcessExist = function (timeout) {
    var result;

    Indicator.Hide();
    result = Sys.WaitProcess(this.getProcessName(), timeout ? timeout : 0).Exists;
    Indicator.Show();

    return result;
  };

  function getRegistryValue(name, defaultValue) {
    var bitPrefix = Sys.OSInfo.Windows64bit ? "Wow6432Node\\" : "";
    var path = aqString.Format("HKEY_LOCAL_MACHINE\\SOFTWARE\\%s%s", bitPrefix, name);
    var result = defaultValue;

    try {
      result = WshShell.RegRead(path);
    }
    catch (ignore) {
    }
    return result;
  }

  this.isInstalled = function () {
    return getRegistryValue("VideoLan\\", "novalue") !== "novalue";
  };

  this.getPath = function () {
    return getRegistryValue("VideoLAN\\VLC\\InstallDir") + "\\" + this.getProcessName() + ".exe";
  };
}

// Predefined video quality settings
function Presets() {
  var _normal = {
    name: "Normal",
    fps: 24,
    quality: 1000
  };
  var _low = {
    name: "Low",
    fps: 20,
    quality: 500
  };
  var _high = {
    name: "High",
    fps: 30,
    quality: 1600
  };
  var _default = _normal;

  this.getDefault = function () {
    return _default;
  };

  this.get = function (name) {
    var presets = [_normal, _low, _high];
    var lowerNameToFind = _default.name.toLowerCase();
    var i, found = _default;

    if (typeof name === "string") {
      lowerNameToFind = name.toLowerCase();
    }

    for (i = 0; i < presets.length; i++) {
      if (presets[i].name.toLowerCase() === lowerNameToFind) {
        found = presets[i];
        break;
      }
    }

    return found;
  };
}

// Creates an object that provides info on the recorded video file
function VideoFile() {
  var _path = (function generateVideoFilePath() {
    var now = aqDateTime.Now();

    var year = aqDateTime.GetYear(now);
    var month = aqDateTime.GetMonth(now);
    var day = aqDateTime.GetDay(now);
    var hour = aqDateTime.GetHours(now);
    var minute = aqDateTime.GetMinutes(now);
    var sec = aqDateTime.GetSeconds(now);

    return Log.Path + "video_" + [year, month, day, hour, minute, sec].join("-") + ".mp4";
  })();

  this.getPath = function () {
    return _path;
  };
}

// Creates an object that provides information on the mouse pointer
function CursorFile() {
  var _path = aqFileSystem.ExpandUNCFileName("%temp%\\vlc_cursor.png");

  if (!aqFile.Exists(_path)) {
    (function createCursorFile(size, color, format, path) {
      var picture = Sys.Desktop.Picture(0, 0, size, size);
      var config = picture.CreatePictureConfiguration(format);
      var i, j;

      config.CompressionLevel = 9;
      for (i = 0; i < size; i++) {
        for (j = 0; j < size; j++) {
          picture.Pixels(i, j) = color;
        }
      }

      picture.SaveToFile(path, config);
    })(12, 0x0000FF /* Red color */, "png", _path);
  }

  this.getPath = function () {
    return _path;
  };
}

// Creates a wrapper object for the VLC recorder
function RecorderEngine() {
  var _recorderInfo = new RecorderInfo();
  var _presets = new Presets();

  var _settings = _presets.getDefault();
  var _videoFile;
  var _cursorFile;
  var _isStarted = false;

  function runCommand(args) {
    WshShell.Run(aqString.Format('"%s" %s', _recorderInfo.getPath(), args), 2, false);
  }

  function getStartCommandArgs() {
    return "--one-instance screen:// -I dummy :screen-fps=" + _settings.fps +
      " :screen-follow-mouse :screen-mouse-image=" + "\"" + _cursorFile.getPath() + "\"" +
      " :no-sound :sout=#transcode{vcodec=h264,vb=" + _settings.quality + ",fps=" + _settings.fps + ",scale=1}" +
      ":std{access=file,dst=\"" + _videoFile.getPath() + "\"}";
  }

  function runStartCommand() {
    runCommand(getStartCommandArgs());
  }

  function runStopCommand() {
    runCommand("--one-instance vlc://quit");
  }

  function ensureRecorderProcessIsClosed(timeout) {
    var timeoutPortion = 1000;
    var process = Sys.WaitProcess(_recorderInfo.getProcessName());
    var wastedTime = 0;
    while (process.Exists) {
      Delay(timeoutPortion, messages.encodingInProgress);
      wastedTime += timeoutPortion;
      if (wastedTime >= timeout) {
        process.Terminate();
        Log.Warning(logMessages.processWasTerminated.message, logMessages.processWasTerminated.messageEx);
      }
    }
  }

  this.getPresetName = function () {
    return _settings.name;
  };

  this.start = function (presetName) {
    var recExists = _recorderInfo.doesProcessExist();

    if (recExists) {
      Log.Warning(logMessages.startFailAlreadyStarted.message, aqString.Format(logMessages.startFailAlreadyStarted.messageEx, _recorderInfo.getProcessName()));
      return "";
    }

    if (!_recorderInfo.isInstalled()) {
      var pmHigher = 300;
      var attr = Log.CreateNewAttributes();
      attr.ExtendedMessageAsPlainText = false;
      Log.Warning(logMessages.recorderIsNotInstalled.message, aqString.Format(logMessages.recorderIsNotInstalled.messageEx, _recorderInfo.getHomepage(), _recorderInfo.getHomepage()), pmHigher, attr);
      return "";
    }

    _settings = _presets.get(presetName);
    _videoFile = new VideoFile();
    _cursorFile = new CursorFile();
    runStartCommand();
    if (!_recorderInfo.doesProcessExist(3000) /* Wait for 3 seconds for the recorder to start */) {
      Log.Warning(logMessages.startNoRecorderProcess.message, aqString.Format(logMessages.startNoRecorderProcess.messageEx, _recorderInfo.getProcessName(), _recorderInfo.getPath(), getStartCommandArgs()));
      return "";
    }

    _isStarted = true;
    Log.Message(logMessages.startOk.message, aqString.Format(logMessages.startOk.messageEx, _settings.name, _videoFile.getPath()));
    return _videoFile.getPath();
  };

  this.stop = function () {
    var recExists = _recorderInfo.doesProcessExist(1000);
    var i;

    if (!recExists) {
      Log.Warning(logMessages.stopFailNoRecorderProcess.message, logMessages.stopFailNoRecorderProcess.messageEx);
      return "";
    }

    if (!_isStarted) {
      Log.Warning(logMessages.stopFailRecorderNotStarted.message, aqString.Format(logMessages.stopFailRecorderNotStarted.messageEx, _recorderInfo.getProcessName()));
      return "";
    }

    Indicator.Hide();
    Delay(2000); // 2 s delay for last frames
    runStopCommand();
    Delay(1000); // 1 s delay to exclude the encoding message from the video
    Indicator.Show();
    Indicator.PushText(messages.encodingInProgress);

    // Terminate the recorder on timeout (if encoding takes too much time)
    Log.Enabled = false;
    ensureRecorderProcessIsClosed(10 * 60 * 1000 /* Wait for 10 minutes for the recorder to encode the video file */);
    Log.Enabled = true;

    _isStarted = false;
    for (i = 0; i < 20; i++) {
      if (aqFile.Exists(_videoFile.getPath())) {
        break;
      }
      Delay(1000, messages.encodingInProgress);
    }

    if (aqFile.Exists(_videoFile.getPath())) {
      Log.Link(_videoFile.getPath(), logMessages.stopOk.message, aqString.Format(logMessages.stopOk.messageEx, _videoFile.getPath()));
    }
    else {
      Log.Warning(logMessages.recorderUnexpectedError.message, aqString.Format(logMessages.recorderUnexpectedError.messageEx, _recorderInfo.getProcessName(), _recorderInfo.getPath(), getStartCommandArgs()));
    }
    return _videoFile.getPath();
  };

  this.isRecording = function () {
    return _isStarted;
  };

  this.onInitialize = function () {
  };

  this.onFinalize = function () {
    if (_isStarted) {
      runStopCommand();
    }
  };
}

// Creates the recorder engine
var gRecorderEngine = new RecorderEngine();

// This method is called on loading the extension
function Initialize() {
  gRecorderEngine.onInitialize();
}

// This method is called on unloading the extension
function Finalize() {
  gRecorderEngine.onFinalize();
}

//
// Runtime scripting object
//

function RuntimeObject_Start(VideoQuality) {
  if (typeof Log === "undefined") { // Check if a test is running or not (to avoid issues with Code Completion)
    return "";
  }

  return gRecorderEngine.start(VideoQuality);
}

function RuntimeObject_Stop() {
  if (typeof Log === "undefined") { // Check if a test is running or not (to avoid issues with Code Completion)
    return "";
  }

  return gRecorderEngine.stop();
}

function RuntimeObject_IsRecording() {
  if (typeof Log === "undefined") { // Check if a test is running or not (to avoid issues with Code Completion)
    return "";
  }

  return gRecorderEngine.isRecording();
}

//
// "Start video recording" keyword-test operation
//

function KDTStartOperation_OnCreate(Data, Parameters) {
  return true;
}

function KDTStartOperation_OnExecute(Data, Parameters) {
  return gRecorderEngine.start();
}

function KDTStartOperation_OnSetup(Data, Parameters) {
  return true;
}

//
// "Stop video recording" keyword-test operation
//

function KDTStopOperation_OnCreate(Data, Parameters) {
  return true;
}

function KDTStopOperation_OnExecute(Data, Parameters) {
  return gRecorderEngine.stop();
}

function KDTStopOperation_OnSetup(Data, Parameters) {
  return true;
}
