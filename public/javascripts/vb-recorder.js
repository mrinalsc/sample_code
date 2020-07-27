//@TODO configure video/audio - bitrate
//@TODO estimate network speed and configure 

var mediaRecorder;
var mediaStreamGlobal;
var getMediaDevicesUntilGetIt;
var mediaTimer;
var progressInSeconds;
var mediaConfiguration = {
    audio : {
        sampleSize: { ideal:16 },
        sampleRate: { ideal: 16000},
        echoCancellation: { ideal: false},
        autoGainControl: { ideal: false},
        noiseSuppression: { ideal: false},
        volume: { ideal: 1.0 }
    },
    video: {
        width: { ideal: 1920 },
        height: { ideal: 1080 },
        aspectRatio : { ideal: 1.7777777778 }
    }
}

function getMediaAccess() {
    return new Promise(function (resolve, reject) {
        navigator.mediaDevices.getUserMedia(mediaConfiguration).then(function (mediaStream) {
            resolve(mediaStream);
        }).catch(function (e) {
            logEvent('Media access denied', { error: e});
            reject(e);
        });
    })
}

function initialiseMediaDevices() {
    logEvent('Initialising Media for user');

    getMediaAccess()
    .then(function (mediaStream) {
        if (getMediaDevicesUntilGetIt) {
            clearInterval(getMediaDevicesUntilGetIt);
        }
        mediaStreamGlobal = mediaStream;
        videoPlayerElement.srcObject = mediaStream;
        videoBoothRenderer.recordingControls.show('startElements');
    })
    .catch(function (e) {
        logEvent('Media recording stopped', {error: e});
        if (mediaAccessErrorDisplayer.classList.contains('w3-hide')) {
            videoBoothRenderer.recordingControls.show('mediaAccessDeniedElements');
        }
    });
}

//AbortError //TrackStartError
function RecordingDone(){
    configSubmissionForm(blobAppendUrl.url.split('?')[0]);
    if(videoBoothConfig.fullUpload) {
        uploadFullFile();
    }
    showSubmissionPage(1)
}

function startRecording() {
    setupProgressUI(videoBoothConfig.recordingTime); // initialise & reset progressbars configurations
    initialiseAzureUploader(); // initialise and create a new azure blob upload url
    
    if(mediaStreamGlobal){
        mediaRecorder = new MediaRecorder(mediaStreamGlobal);
        mediaRecorder.start(500);
        setupDynamicInstructionRendering();// initialise and create setTimeouts to render custom instructions
        videoBoothRenderer.recordingControls.show('inProgressElements');
        progressInSeconds = 1;
        
        logEvent('Media recording started');

        mediaRecorder.ondataavailable = function (media) {
            updateRecordingProgress(progressInSeconds);
            if (media.data.size > 0) {
                uploadABlobToAzure(media.data);
            }
            progressInSeconds += 0.5;
        }

        mediaTimer = setTimeout(function () {
            if(mediaRecorder.state === 'recording') {
                mediaRecorder.stop();
                setupProgressUI(videoBoothConfig.recordingTime);
                videoBoothRenderer.recordingControls.show('uploadingElements');
                logEvent('Media recording stopped');
            }
            else{
                progressInSeconds = 0;
                setupProgressUI(videoBoothConfig.recordinguploadWaitTime);
                configSubmissionForm(blobAppendUrl.url.split('?')[0]);
                showSubmissionPage(videoBoothConfig.recordinguploadWaitTime);
            }
        }.bind(this), videoBoothConfig.recordingTime * 1000);

        mediaRecorder.onstop = RecordingDone;
    } else {
        checkMediaAccessPermissionRegularly();
    }
}

function showSubmissionPage(waitTime){
    var checkForUploadCompletion = setInterval(function(){
        //if(totalFullFileUploadedBytes + totalPartialFileUploadedBytes === getTotalFileSize()
        //|| progressInSeconds >= videoBoothConfig.recordinguploadWaitTime) {

        if( isAllDataUploaded() 
            || progressInSeconds >= videoBoothConfig.recordinguploadWaitTime) {
            clearInterval(checkForUploadCompletion);
            videoBoothRenderer.recordingControls.show('completedElements');
        }
        else{
            updateUploadingProgress(progressInSeconds, videoBoothConfig.recordinguploadWaitTime);
            progressInSeconds = progressInSeconds + 1;
        }
    }, waitTime*1000);
}

function isAllDataUploaded() {
 return totalFullFileUploadedBytes + totalPartialFileUploadedBytes === getTotalFileSize()
}

function isUploadfailedAfterMinimumTimeRequired() {
    return blockUploadFailedCount;
}

function stopMedia() {
    if(mediaRecorder.state === 'recording') {
        mediaRecorder.stop();
    }

    videoPlayerElement.srcObject = null;
    mediaStreamGlobal.getTracks().forEach(function(track) {
        track.stop();
    });
}

function initialiseRecordingPage() {
    // this will handle if we have camera permission on page load
    initialiseMediaDevices();
    checkMediaAccessPermissionRegularly();
    setupProgressUI(videoBoothConfig.recordingTime);
    keepVideoPlayerAspectRatio();
}

// this will handle the scenario if we don't have camera permission on page load
function checkMediaAccessPermissionRegularly() {
    getMediaDevicesUntilGetIt = setInterval(function () {
        if (!mediaStreamGlobal) {
            console.log('initialising again');
            initialiseMediaDevices();
        }
    }.bind(this), 2000);
}


function checkMediaAccessPermissionRegularly() {
    getMediaDevicesUntilGetIt = setInterval(function () {
        if (!mediaStreamGlobal) {
            console.log('initialising again');
            initialiseMediaDevices();
        }
    }.bind(this), 2000);
}



