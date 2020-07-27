function validateWorker(workerId) {
    var validateWorkerRequest = new XMLHttpRequest();
    validateWorkerRequest.open('GET', '/api/workers/' + workerId, false);
    validateWorkerRequest.setRequestHeader('Content-Type', 'application/json');
    
    try {
        validateWorkerRequest.send();
        if(validateWorkerRequest.status === 200) {
            logEvent('New worker allowed');
            return true;
        } else {
            logEvent('Duplicate worker tried to do hit again');
            return false;
        }
    } catch(err) {
        logEvent('Duplicate worker validation failed' ,{ error: e});
        return false;
    }
}


function saveWorkerId(workerId) {
    var saveWorkerIdRequest = new XMLHttpRequest();
    saveWorkerIdRequest.open('POST', '/api/workers', true);
    saveWorkerIdRequest.setRequestHeader('Content-Type', 'application/json');
    try {
        saveWorkerIdRequest.send(JSON.stringify({ "workerID" : workerId }));
    } catch(err) {
        logEvent('Worker Id failed to save in azure db', {error: err});
    }
}
