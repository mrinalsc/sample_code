/**Video Configurations - customizable
 * recordingTime = values in seconds
 * customInstruction = keys should the seconds(in integer) in which the associated values need to be shown,
 *                     values need to be text/string in Array, Each text/string will be added as paragraph.
 * getStandardInstruction = populate standard instruction with recording time, need to be string.
 * **/
var videoBoothConfig = {
    recordingTime: 35, //In seconds
    customInstructions: { 
        //0 : ["<p style='color:red;'>Please read the following text:</p><p>The woman led the children still deeper into the forest, where they had never in their lives been before. Then a great fire was again made, and the mother said: ‘Just sit there, you children, and when you are tired you may sleep a little; we are going into the forest to cut wood, and in the evening when we are done, we will come and fetch you away.’ When it was noon, Gretel shared her piece of bread with Hansel, who had scattered his by the way. Then they fell asleep and evening passed, but no one came to the poor children.</p>"]
        // 5 : ["After 5 seconds it will show"],
        // 7 : ["After 7 seconds this will be shown"],
        // 9 : ["This is something we will show at 9 seconds until the end of 10 sec"]
    },
    getStandardInstructions: function() {
        return []
    },
    partialUpload : true,
    fullUpload : false, 
    recordinguploadWaitTime: 180 //In seconds
}
var assignmentIdInPreviewMode = 'ASSIGNMENT_ID_NOT_AVAILABLE';

function getConfigUri() {
    var url = window.location.href;
    var regex = new RegExp('[?]conf=([A-z]*)');
    var results = regex.exec(url);
    return results === null ? '' : results[0];
};


function getAssignmentId() {
    var assignmentId = getParam("assignmentId");
    return assignmentId || 'development';
}

function getParam(param) {
    var searchingParam = document.location.search;
    searchingParam = searchingParam.substr(searchingParam.indexOf(param) + param.length + 1).split('&')[0];
    return searchingParam || 'development';
}

function isBrowserCompatible() {
    try { 
        MediaRecorder
    }
    catch(error) { 
        logEvent("Incompatible browser", {error: err});
        return false
    }
    return !!MediaRecorder && navigator.mediaDevices && !!window.fetch && window.navigator.userAgent.indexOf('MSIE') === -1
}

function isNewWorker() {
    // Verify if worker already completed the hit
    return true;

}

function isMacOs() {
    return !!navigator.platform && /iPad|iPhone|iPod/.test(navigator.platform);
}

function getContainerToRender() {
    if(isMacOs()) { return 'incompatibleOSContainer'};
    if(!isBrowserCompatible()) { return 'incompatibleBrowserContainer' };
    return isNewWorker() ? 'consentFormContainer' : 'existingWorkerContainer';
}

function showConsentForm() {
    if(assignmentId !== assignmentIdInPreviewMode) {
        setTimeout(function(){
            var containerToShow = getContainerToRender();
            videoBoothRenderer.show(containerToShow);
            videoBoothRenderer.hide('welcomeContainer');
            logEvent(containerToShow + ' shown to user');
        }, 0);
    } else {
            logEvent('Consent form denied to user in preview mode')
    }
}

function showRecordingPage(e) {
    e.preventDefault();
    e.stopImmediatePropagation();
    if (validateInputs('worker-info-input')) {  
        videoBoothRenderer.hide('workerInfoForm');
        videoBoothRenderer.hide('completedContainer');
        videoBoothRenderer.show('recordingContainer');
        initialiseRecordingPage();
        logEvent('Shown recording page');
    }
}

function showWorkerForm(e) {
    e.preventDefault();
    e.stopImmediatePropagation();
    if (validateInputs('vb-consent-input')) {
 
        videoBoothRenderer.hide('consentFormContainer');
        videoBoothRenderer.show('workerInfoForm');
        logEvent('Shown worker info form');
        const Http = new XMLHttpRequest();
        const url=document.URL+"&activityStatus=CONSENT";
        Http.open("GET", url);
        Http.send();
        
        
    }
    else { 
        function getQueryStringValue (key) {  
            return decodeURIComponent(window.location.search.replace(new RegExp("^(?:.*[&\\?]" + encodeURIComponent(key).replace(/[\.\+\*]/g, "\\$&") + "(?:\\=([^&]*))?)?.*$", "i"), "$1"));  
          }; 
           
          workerID=(getQueryStringValue("workerId"));
        window.alert("Dear "+workerID+", Please agree to the Terms and Conditions");}
}

function cancelRecording() {
    videoBoothRenderer.recordingControls.show('startElements');
    mediaRecorder.stop();
    clearTimeout(mediaTimer);
    clearSetTimeOutsToRenderCustomInstructions();
    abortAllUploads();
    deleteUploadUrl();
    logEvent('User cancelled recording');
}

function configSubmissionForm(azureUrl){
    var workerGender = document.getElementById('worker-gender').value;
    var workerAge = document.getElementById('worker-age').value;
    var mturkFormActionUrl = decodeURIComponent(getParam('turkSubmitTo')) + "/mturk/externalSubmit?assignmentId=" + getAssignmentId() + "&gender=" + workerGender + "&age=" + workerAge + "&azureUrl=" + azureUrl;
    mturkForm.action = mturkFormActionUrl;
    logEvent('Configuring user info in submission form : ', {gender: workerGender, age: workerAge });
}


function submitAmazonHit(e) {
    e.preventDefault();
    e.stopImmediatePropagation();
    try {
        logEvent('Submit form:');
        mturkForm.submit();
    } catch(error) {
        logEvent('Mturk Submission Failed', {error: error});
    }
    
}

function showSubmissionForm(e) {
    e.preventDefault();
    e.stopImmediatePropagation();
    stopMedia();
    videoBoothRenderer.hide('recordingContainer');
    videoBoothRenderer.show('completedContainer');
    saveWorkerId(getParam("workerId"));

    logEvent('Submission page shown to user');
}

function logEvent(event, eventData) {
    if (!!appInsights) {
        appInsights.trackEvent(event, eventData );
    }
}

function logUnhandledErrors(error) {
    if (!!appInsights) {
        appInsights.trackEvent('Global Error :', error);
    }
}

window.addEventListener('error', logUnhandledErrors);

document.getElementById("welcome-btn").onclick = showConsentForm;

document.getElementById('consent-agreement-btn').onclick = showWorkerForm;

document.getElementById('worker-info-btn').onclick = showRecordingPage;

document.getElementById('start-recording-btn').onclick = startRecording;

document.getElementById('cancel-recording-btn').onclick = cancelRecording;

document.getElementById('restart-recording-btn').onclick = startRecording;

document.getElementById('submit-task-btn').onclick = submitAmazonHit;

document.getElementById('show-sub-form-btn').onclick = showSubmissionForm;