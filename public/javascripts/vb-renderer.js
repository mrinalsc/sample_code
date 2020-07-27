// Main containers full page displaying Elements
var consentForm = document.getElementById('consent-form-container');
var existingWorkerContainer = document.getElementById('existing-worker-container');
var incompatibleBrowserContainer = document.getElementById('incompatible-browser-container');
var welcomeContainer = document.getElementById('welcome-container');
var recordingContainer = document.getElementById('recording-container');
var videoPlayerContainer = document.getElementById('videoContainerBlock');
var recordingInstructionsContainer = document.getElementById('recording-instructions');
var mturkForm = document.getElementById('mturk-submission');
var videoPlayerElement = document.getElementById('videoPlayer');
var completedContainer = document.getElementById('completed-container');
var incompatibleOSContainer = document.getElementById('incompatible-os-container');

// Recording control elements
var startButton = document.getElementById('start-recording-btn');
var startInstruction = document.getElementById('start-instruction');
var cancelButton = document.getElementById('cancel-recording-btn');
var progressBar = document.getElementById('recording-progress-bar');
var uploadingProgressBar = document.getElementById('uploading-progress-bar');
var timeDisplayer = document.getElementById('recording-time-displayer');
var mediaAccessErrorDisplayer = document.getElementById('allow-media-error');
var restartRecordingBtn = document.getElementById('restart-recording-btn');
var nextBtnToShowSubForm = document.getElementById('show-sub-form-btn');
// var restartBtnInRecordingPage = document.getElementById('restart-recording-btn-in-rec-page');
var completedDisplayer = document.getElementById('job-completed');
var uploadingDisplayer = document.getElementById('alert-uploading');
var uploadFailedDisplayer = document.getElementById('upload-failure');
var workerInfoForm = document.getElementById('worker-info-container');
var submitBtn = document.getElementById('submit-task-btn');
var waitingForMediaInstruction = document.getElementById('wait-for-media');
var timerElement = document.getElementById('recording-time-left');

var dynamicInstructionContainer = document.getElementById('dynamic-instructions');

var setTimeOutsToRenderCustomInstructions = [];

var videoBoothRenderer = {
    htmlContainers: {
        'consentFormContainer': consentForm,
        'existingWorkerContainer': existingWorkerContainer,
        'workerInfoForm': workerInfoForm,
        'incompatibleBrowserContainer': incompatibleBrowserContainer,
        'incompatibleOSContainer' : incompatibleOSContainer,
        'welcomeContainer': welcomeContainer,
        'recordingContainer': recordingContainer,
        'completedContainer': completedContainer
    },

    show: function (element) {
        this.htmlContainers[element].classList.remove('w3-hide');
        if(!!appInsights) {
            appInsights.trackPageView(element);
        }
    },

    hide: function (element) {
        this.htmlContainers[element].classList.add('w3-hide');
    },

    recordingControls: {
        controlElements: {
            'initElements': [waitingForMediaInstruction],
            'startElements': [startButton, startInstruction],
            'inProgressElements': [progressBar, timeDisplayer, cancelButton],
            'mediaAccessDeniedElements': [mediaAccessErrorDisplayer],
            'completedElements': [restartRecordingBtn, completedDisplayer, nextBtnToShowSubForm],
            'uploadFailedElements': [uploadFailedDisplayer],
            'uploadingElements': [uploadingDisplayer, uploadingProgressBar],
            'allElements': [startButton, startInstruction, progressBar, timeDisplayer, cancelButton,
                mediaAccessErrorDisplayer, uploadingProgressBar, restartRecordingBtn, completedDisplayer,
                uploadFailedDisplayer, uploadingDisplayer, waitingForMediaInstruction, nextBtnToShowSubForm]
        },

        show: function (controlsToShow) {
            var allElements = this.controlElements.allElements;
            var elementsToShow = this.controlElements[controlsToShow];

            allElements.forEach(function (currentElement) {
                if (elementsToShow.indexOf(currentElement) > -1){
                    currentElement.classList.remove('w3-hide');
                } else if (!currentElement.classList.contains('w3-hide')) {
                        currentElement.classList.add('w3-hide');
                }
            });
        }
    }
}

function updateProgress(progressBar, timerElement ) {
    return function (recordingTime) {
        var recordingTime = recordingTime < 0 ? 0 : recordingTime;
        var progressBarWidth = (videoBoothConfig.recordingTime - recordingTime) / videoBoothConfig.recordingTime * 100;

        if (recordingTime > 0) {
            progressBar.style.width = progressBarWidth + "%";
            if(timerElement) {
                timerElement.innerHTML = videoBoothConfig.recordingTime - Math.floor(recordingTime);
            }
        }
        else {
            progressBar.style.width = "0%";
            if(timerElement) {
                timerElement.innerHTML = '';
            }
        }
    }
}

var updateRecordingProgress = updateProgress(progressBar, timerElement);

function updateUploadingProgress(currentTime, totalTime){

    if(currentTime && totalTime){
        uploadingProgressBar.style.width = Math.round((currentTime*100)/totalTime) + "%";
        return;
    }
    var totalSizeForProgressBar = getTotalFileSize();
    var progressBarWidth = (totalSizeForProgressBar - (totalFullFileUploadedBytes + totalPartialFileUploadedBytes)) / totalSizeForProgressBar * 100;
    uploadingProgressBar.style.width = progressBarWidth + "%";
}

function renderStaticInstructions() {
    var instructionContainer = document.getElementById('static-instructions');
    videoBoothConfig.getStandardInstructions().forEach(function(instruction) {
        createAndAppendElement(instructionContainer, instruction);
    });
}

function createAndAppendElement(container, content) {
    var contentElement = document.createElement('p');
    contentElement.innerHTML = content;
    container.appendChild(contentElement);
}

function emptyContainer(container) {
    container.innerHTML = '';
}

function validateInputs(form) {  // refactor before handover
    var inputElements = document.getElementsByClassName(form);
    var validForm = false;
    for (var i = 0; i < inputElements.length; i++) {
        validForm = true;

        var currentElement = inputElements[i];
        if (!currentElement.checkValidity()) {

            if (currentElement.type === "checkbox") {
                currentElement.parentElement.classList.add('vb-invalid');
            } else {
                currentElement.classList.add('vb-invalid');
            }
            validForm = false;

        } else {
            if (currentElement.type === "checkbox") {
                currentElement.parentElement.classList.remove('vb-invalid');
            } else {
                currentElement.classList.remove('vb-invalid');
            }
        }
    }

    return validForm;
}

//@TODO Initialise all requireds in a separate initi fn 
function fillDateinConsentForm() {
    document.getElementById('agreementDate').valueAsNumber = Number(new Date()) - ((new Date()).getTimezoneOffset() * 60000);
}

function keepVideoPlayerAspectRatio() {
    var player = document.getElementsByTagName('video')[0];
    player.width = welcomeContainer.offsetWidth * 2/3;
}

function adjustInstructionsContainer(){
    if(recordingInstructionsContainer.offsetHeight > videoPlayerContainer.offsetHeight ) {
        recordingInstructionsContainer.classList.add('vb-v-scrollable');
    }
}

function setupProgressUI(waitTime) {
    timerElement.innerHTML = waitTime;
    progressBar.style.width = '100%';
    uploadingProgressBar.style.width = '100%';
}

function setupDynamicInstructionRendering() {
    Object.keys(videoBoothConfig.customInstructions).forEach(function(timeToDisplay){
        var currentSetTimeOut = setTimeout(renderCustomInstruction.bind(this, [timeToDisplay]), Number(timeToDisplay) * 1000);
        setTimeOutsToRenderCustomInstructions.push(currentSetTimeOut);
    });
}

function renderCustomInstruction(timeToDisplay){
    emptyContainer(dynamicInstructionContainer);
    videoBoothConfig.customInstructions[timeToDisplay].forEach(function(instruction){
        createAndAppendElement(dynamicInstructionContainer, instruction);
    });
}

function clearSetTimeOutsToRenderCustomInstructions() {
    setTimeOutsToRenderCustomInstructions.forEach(function(currentSetTimeOut){
        clearTimeout(currentSetTimeOut);
    });

    setTimeOutsToRenderCustomInstructions = [];
}



renderStaticInstructions();
fillDateinConsentForm();
