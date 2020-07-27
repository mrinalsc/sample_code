
		//function getQueryStringValue (key) {  
        //   return decodeURIComponent(window.location.search.replace(new RegExp("^(?:.*[&\\?]" + encodeURIComponent(key).replace(/[\.\+\*]/g, "\\$&") + "(?:\\=([^&]*))?)?.*$", "i"), "$1"));  
        //  }; 
           
     //    workerID=(getQueryStringValue("workerId"));
     workerID='msc';
     val='CONSENT';
          var sql = require('mssql');
          var d = Date();
		var current_hour = d.toString();
            //2.
            var config = {
                server: 'logging.database.windows.net',
                database: 'audit',
                user: 'pprr756',
                password: 'pprr1122#',
                port: 1433
            };
            
           
                //2.
                var dbConn = new sql.ConnectionPool(config);
                //3.
                dbConn.connect().then(function () {
                  //4.
                  var transaction = new sql.Transaction(dbConn);
                  //5.
                  transaction.begin().then(function () {
                      //6.
                      var request = new sql.Request(transaction);
                      //7.
                      var myArgs = process.argv.slice(2);
                      request.query("Insert into dbo.audio_audit (workerID,Completion_Stage,EventTimeStamp) values ("+"'"+workerID+"'"+","+"'"+val+"'"+","+"'"+current_hour+"'"+")")
                  .then(function () {
                          //8.
                          transaction.commit().then(function (recordSet) {
                              console.log(recordSet);
                              dbConn.close();
                          }).catch(function (err) {
                              //9.
                              console.log("Error in Transaction Commit " + err);
                              dbConn.close()
                          });
                      }).catch(function (err) {
                          //10.
                          console.log("Error in Transaction Begin " + err);
                          dbConn.close();
                      })
                       
                  }).catch(function (err) {
                      //11.
                      console.log(err);
                      dbConn.close();
                  })
              }).catch(function (err) {
                  //12.
                  console.log(err);
              })
