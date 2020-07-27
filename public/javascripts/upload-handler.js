//initialise config asyncly in a separate init fn
var sasString = "?sv=2019-02-02&ss=bfqt&srt=sco&sp=rwdlacup&se=2020-02-07T02:20:19Z&st=2019-12-13T18:20:19Z&spr=https&sig=j7fFtOKOwVhCSQGNLzOdHn9b%2BZf5b37xIyMmgx81UOs%3D";
var accountName = "newms";
var containerName = 'recordings';

var containerURL = new azblob.ContainerURL("https://" + accountName + ".blob.core.windows.net/" + containerName + "?" + sasString,
    azblob.StorageURL.newPipeline(new azblob.AnonymousCredential(), {
        retryOptions: { maxTries: 20 }
    }));
var assignmentId = getAssignmentId();
var randomFileName;
var fileNameIncrementer = 0;
var blobURL;
var blobAppendUrl;
var blockBlobUrlForFullFile;
var blockUploadedCount = 0;
var blockUploadFailedCount = 0;
var uploadAborters;
var videoChunks = [];
var totalFileSize = 0;
var totalFullFileUploadedBytes = 0;
var totalPartialFileUploadedBytes = 0;


function generateAppendUrls() {
    return new Promise(function(resolve, reject){
        try {
        ++fileNameIncrementer;
        randomFileName = getRandomNumberFromWorkerId() + String(Number(new Date())) + getRandomNumberFromWorkerId() + fileNameIncrementer;
        [blobAppendUrl, blockBlobUrlForFullFile] = createAppendUrlInAzure(randomFileName);
        resolve();
        } catch(err) {
            logEvent("Error generating the url", {error: err});
            reject();
        }
    });
}

function createAppendUrlInAzure(randomFileName) {
    var blobURL = azblob.BlobURL.fromContainerURL(containerURL, randomFileName + ".webm");
    var blockBlobUrlForFullFile = azblob.BlockBlobURL.fromContainerURL(containerURL, randomFileName + "-full.webm");
    var blobAppendUrl = azblob.AppendBlobURL.fromBlobURL(blobURL);
    
    blobAppendUrl.create(azblob.Aborter.none);
    logEvent("Video generated : "+randomFileName); //TODO:REMOVE. HERE JUST FOR A FEW TEST RUNS

    return [blobAppendUrl, blockBlobUrlForFullFile  ];
}

function initialiseAzureUploader() {
    blobAppendUrl = undefined;
    blockBlobUrlForFullFile = undefined;
    uploadAborters = [];
    blockUploadedCount = 0;
    blockUploadFailedCount = 0;
    totalFileSize = 0;
    totalFullFileUploadedBytes = 0;
    totalPartialFileUploadedBytes = 0;
}

function abortAllUploads() {
    uploadAborters.forEach(function(aborter){
        aborter.abort();
    });
    logEvent('Upload aborted for user');
}

function getRandomNumberFromWorkerId() {
    return assignmentId.charCodeAt(Math.floor(Math.random() * assignmentId.length));
}

//@TODO handle packet loss/failure during azure upload
async function uploadABlobToAzure(data) {
    var currentAborter = new azblob.Aborter();
    logEvent('Uploading data for user');
    uploadAborters.push(currentAborter);
    videoChunks.push(data);
    totalFileSize += data.size;

    if(!blobAppendUrl && !blockBlobUrlForFullFile ) {
        this.chunkUploader = generateAppendUrls();   
    }
    if(videoBoothConfig.partialUpload) {
        this.chunkUploader = this.chunkUploader.then(async function(){
            await blobAppendUrl.appendBlock(
                currentAborter,
                data,
                data.size,
                {
                    blockSize: 1024 * 512,
                    progress: function (progress) {
                    }
                }).then(function(data){
                    totalPartialFileUploadedBytes += data._response.request.body.size;
                    updateUploadingProgress();
                }).catch(function (e) {
                    logEvent('Upload failed for user', { error: e});
                //     if(e.code !== 'REQUEST_ABORTED_ERROR' && e.statusCode !== 404) {
                //         //when upload fails for first time, we will stop recording and show failed message to user
                //         if(!blockUploadFailedCount && progressInSeconds < videoBoothConfig.minimumTimeForPay) {
                //             if(mediaRecorder.state === 'recording') {
                //                 mediaRecorder.stop();
                //             }
                //             clearTimeout(mediaTimer);
                //             abortAllUploads();
                //             videoBoothRenderer.recordingControls.show('uploadFailedElements');
                //         }
                //         ++blockUploadFailedCount;
                //     }
                 });
        });
    }
}

function uploadFullFile() {
    var fullRecording = new Blob(videoChunks, { type: 'video/webm' });
    var currentAborter = new azblob.Aborter();
    logEvent('Uploading full recording for user');
    uploadAborters.push(currentAborter);
    blockBlobUrlForFullFile.upload(currentAborter,
                                    fullRecording,
                                    fullRecording.size, 
                                    { progress: function(progress) {
                                        totalFullFileUploadedBytes = progress.loadedBytes;
                                        updateUploadingProgress();
            }})
        .then(function(d){
            logEvent('Full Upload success for user');
        }).catch(function(e){
            logEvent('Full Upload failed for user', { error: e });
        })
}

function deleteUploadUrl() {
    if(blobURL) {
        blobURL.delete();
        logEvent('Azure file deleted for user', {blob: blobURL });
    }
}

function getTotalFileSize() {
    return (videoBoothConfig.partialUpload ? totalFileSize : 0) + (videoBoothConfig.fullUpload ? totalFileSize : 0);
}