/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office */



Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;

    //$('#get-slide-metadata').click(getSlideMetadata);  
    document.getElementById("input-localport").onchange = setupReceiver;

    setupReceiver();    

   /* if(server) server.close();

   var WebSocketServer = require('websocket').server;
    var http = require('http');
    var server = http.createServer(function(request, response) {
        console.log((new Date()) + ' Received request for ' + request.url);
        response.writeHead(404);
        response.end();
    });

    server.listen(8080, function() {
        console.log((new Date()) + ' Server is listening on port 8080 now');
    });
    */
  } 
}); 

 
export async function setupReceiver()
{
  console.log("Binding on port "+localPort);
  var localPort =  document.getElementById("input-localport").value;
  const WebSocket = require('ws');
  const wss = new WebSocket.Server({
  port: localPort,
  perMessageDeflate: {
    zlibDeflateOptions: {
      // See zlib defaults.
      chunkSize: 1024,
      memLevel: 7,
      level: 3
    }, 
    zlibInflateOptions: {
      chunkSize: 10 * 1024
    },
    // Other options settable:
    clientNoContextTakeover: true, // Defaults to negotiated value.
    serverNoContextTakeover: true, // Defaults to negotiated value.
    serverMaxWindowBits: 10, // Defaults to negotiated value.
    // Below options specified as default values.
    concurrencyLimit: 10, // Limits zlib concurrency for perf.
    threshold: 1024 // Size (in bytes) below which messages
    // should not be compressed.
  }
});
}
 
/*
export async function getSlideMetadata() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
      function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              showNotification("Error", asyncResult.error.message);
          } else {
              showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
          }
      }
  );
}
*/ 