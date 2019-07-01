// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.

// Require jQuery
const $ = jQuery= require('jquery');

// Require powershell
const powershell = require('node-powershell');

// Require datatables
const dt = require('datatables.net')();
const dtbs = require('datatables.net-bs4')(window, $);
const dtrs = require( 'datatables.net-responsive-dt')( window, $ );
const dtrg = require( 'datatables.net-rowgroup-dt')( window, $ );
const Spinner = require('spin');
const tippy = require('tippy.js');
const mom = require('moment');
var moment = mom;
moment.locale('ru');

// Set dates 
let startdatetime = new Date(Date.now() - 3600000  - (new Date().getTimezoneOffset()*1000*60)).toISOString().substring(0,16);
document.querySelector("#StartTime-form").value = startdatetime;

let enddatetime = new Date(Date.now() - (new Date().getTimezoneOffset()*1000*60)).toISOString().substring(0,16);
document.querySelector("#EndTime-form").value = enddatetime;


// TOOLTIP SOURCE SECTION
sourcetip = {
    ADMIN:'The event source was human intervention. For example, an administrator used Queue Viewer to delete a message, or submitted message files using the Replay directory.',
    AGENT:'The event source was a transport agent.',
    APPROVAL:'The event source was the approval framework thats used with moderated recipients. For more information, see Moderated Transport.',
    BOOTLOADER:'The event source was unprocessed messages that exist on the server at boot time. This is related to the LOAD event type.',
    DNS:'The event source was DNS.',
    DSN:'The event source was a delivery status notification (also known as a DSN, bounce message, non-delivery report, or NDR).',
    GATEWAY:'The event source was a Foreign connector. For more information, see Foreign Connectors.',
    MAILBOXRULE:'The event source was an Inbox rule. For more information, see Inbox rules.',
    MEETINGMESSAGEPROCESSOR:'The event source was the meeting message processor, which updates calendars based on meeting updates.',
    ORAR:'The event source was an Originator Requested Alternate Recipient (ORAR). You can enable or disable support for ORAR on Receive connectors using the OrarEnabled parameter on the New-ReceiveConnector or Set-ReceiveConnector cmdlets.',
    PICKUP:'The event source was the Pickup directory. For more information, see Pickup Directory and Replay Directory.',
    POISONMESSAGE:'The event source was the poison message identifier. For more information about poison messages and the poison message queue, see Queues and messages in queues',
    PUBLICFOLDER:'The event source was a mail-enabled public folder.',
    QUEUE:'The event source was a queue.',
    REDUNDANCY:'The event source was Shadow Redundancy. For more information, see Shadow redundancy in Exchange Server.',
    ROUTING:'The event source was the routing resolution component of the categorizer in the Transport service.',
    SAFETYNET:'The event source was Safety Net. For more information, see Safety Net in Exchange Server.',
    SMTP:'The message was submitted by the SMTP send or SMTP receive component of the transport service.',
    STOREDRIVER:'The event source was a MAPI submission from a mailbox on the local server.'
}

eventtip = {
    AGENTINFO:'This event is used by transport agents to log custom data.',
    BADMAIL:'A message submitted by the Pickup directory or the Replay directory that cant be delivered or returned.',
    CLIENTSUBMISSION:'A message was submitted from the Outbox of a mailbox.',
    DEFER:'Message delivery was delayed.',
    DELIVER:'A message was delivered to a local mailbox.',
    DELIVERFAIL:'An agent tried to deliver the message to a folder that doesnt exist in the mailbox.',
    DROP:'A message was dropped without a delivery status notification (also known as a DSN, bounce message, non-delivery report, or NDR).',
    DSN:'A delivery status notification (DSN) was generated.',
    DUPLICATEDELIVER:'A duplicate message was delivered to the recipient. Duplication may occur if a recipient is a member of multiple nested distribution groups. Duplicate messages are detected and removed by the information store.',
    DUPLICATEEXPAND:'During the expansion of the distribution group, a duplicate recipient was detected.',
    DUPLICATEREDIRECT:'An alternate recipient for the message was already a recipient.',
    EXPAND:'A distribution group was expanded.',
    FAIL:'Message delivery failed. Sources include SMTP, DNS, QUEUE, and ROUTING.',
    HADISCARD:'A shadow message was discarded after the primary copy was delivered to the next hop. For more information, see Shadow redundancy in Exchange Server.',
    HARECEIVE:'A shadow message was received by the server in the local database availability group (DAG) or Active Directory site.',
    HAREDIRECT:'A shadow message was created.',
    HAREDIRECTFAIL:'A shadow message failed to be created. The details are stored in the source-context field.',
    INITMESSAGECREATED:'A message was sent to a moderated recipient, so the message was sent to the arbitration mailbox for approval. For more information, see Moderated Transport.',
    LOAD:'A message was successfully loaded at boot.',
    MODERATIONEXPIRE:'A moderator for a moderated recipient never approved or rejected the message, so the message expired. For more information about moderated recipients, see Moderated Transport.',
    MODERATORAPPROVE:'A moderator for a moderated recipient approved the message, so the message was delivered to the moderated recipient.',
    MODERATORREJECT:'A moderator for a moderated recipient rejected the message, so the message wasnt delivered to the moderated recipient.',
    MODERATORSALLNDR:'All approval requests sent to all moderators of a moderated recipient were undeliverable, and resulted in non-delivery reports (also known as NDRs or bounce messages).',
    NOTIFYMAPI:'A message was detected in the Outbox of a mailbox on the local server.',
    NOTIFYSHADOW:'A message was detected in the Outbox of a mailbox on the local server, and a shadow copy of the message needs to be created.',
    POISONMESSAGE:'A message was put in the poison message queue or removed from the poison message queue.',
    PROCESS:'The message was successfully processed.',
    PROCESSMEETINGMESSAGE:'A meeting message was processed by the Mailbox Transport Delivery service.',
    RECEIVE:'A message was received by the SMTP receive component of the transport service or from the Pickup or Replay directories (source: SMTP), or a message was submittedfrom a mailbox to the Mailbox Transport Submission service (source: STOREDRIVER).',
    REDIRECT:'A message was redirected to an alternative recipient after an Active Directory lookup.',
    RESOLVE:'A messages recipients were resolved to a different email address after an Active Directory lookup.',
    RESUBMIT:'A message was automatically resubmitted from Safety Net. For more information, see Safety Net in Exchange Server.',
    RESUBMITDEFER:'A message resubmitted from Safety Net was deferred.',
    RESUBMITFAIL:'A message resubmitted from Safety Net failed.',
    SEND:'A message was sent by SMTP between transport services.',
    SUBMIT:'The Mailbox Transport Submission service successfully transmitted the message to the Transport service.',
    SUBMITDEFER:'The message transmission from the Mailbox Transport Submission service to the Transport service was deferred.',
    SUBMITFAIL:'The message transmission from the Mailbox Transport Submission service to the Transport service failed.',
    SUPPRESSED:'The message transmission was suppressed.',
    THROTTLE:'The message was throttled.',
    TRANSFER:'Recipients were moved to a forked message because of content conversion, message recipient limits, or agents. Sources include ROUTING or QUEUE.'
}



// SPINNER

var opts = {
  lines: 13, // The number of lines to draw
  length: 38, // The length of each line
  width: 17, // The line thickness
  radius: 45, // The radius of the inner circle
  scale: 1, // Scales overall size of the spinner
  corners: 1, // Corner roundness (0..1)
  color: '#ffffff', // CSS color or array of colors
  fadeColor: 'transparent', // CSS color or array of colors
  speed: 1, // Rounds per second
  rotate: 0, // The rotation offset
  animation: 'spinner-line-fade-quick', // The CSS animation name for the lines
  direction: 1, // 1: clockwise, -1: counterclockwise
  zIndex: 2e9, // The z-index (defaults to 2000000000)
  className: 'spinner', // The CSS class to assign to the spinner
  top: '50%', // Top position relative to parent
  left: '50%', // Left position relative to parent
  shadow: '0 0 1px transparent', // Box-shadow for the lines
  position: 'relative' // Element positioning
};

var target = document.getElementById('spinner');
var spinner = new Spinner(opts).spin(target);


// Form submit event
$('#EmailForm').submit(() => {
    event.preventDefault();

// Variables
let Powershellpath = $('#Computername-form').val();
let Server = $('#Server-form').val();
let Subject = $('#Subject-form').val();    
let Sender = $('#Sender-form').val();
let Recipients = $('#Recipients-form').val();
let MessageID = $('#MessageID-form').val();
let HideHealthMessages = $('#ServiceMessages-form').val();
let Source = $('#Source-form').val();
let EventId = $('#EventId-form').val();
let Start = $('#StartTime-form').val(); 
let End = $('#EndTime-form').val();
let Username = $('#Username-form').val();
let Password = $('#Password-form').val();

//console.log(EventIds)

document.getElementById('FormContainer').classList.add('d-none')
document.getElementById('spinner').classList.remove('d-none')
    
 // Create the PS Instance
 let ps = new powershell({
    executionPolicy: 'Bypass',
    noProfile: true
})

ps.addCommand("./get-ExchLogs.ps1");

// Parameters
if (Powershellpath){ps.addParameter({name: 'Powershellpath', value: Powershellpath})};
if (Server){ps.addParameter({name: 'Server', value: Server})};
if (Subject){ps.addParameter({name: 'MessageSubject', value: Subject})};
if (Sender){ps.addParameter({name: 'Sender', value: Sender})};
if (Recipients){ps.addParameter({name: 'Recipients', value: Recipients})};
if (MessageID){ps.addParameter({name: 'MessageID', value: MessageID})};
if (HideHealthMessages =="Yes"){ps.addParameter({name: 'HideHealthMessages', value: HideHealthMessages})};
if (Source){ps.addParameter({name: 'Source', value: Source})};
if (EventId){ps.addParameter({name: 'EventId', value: EventId})};
if (Start){ps.addParameter({name: 'Start', value: Start})};
if (End){ps.addParameter({name: 'End', value: End})};
if (Username){ps.addParameter({name: 'Username', value: Username})};
if (Password){ps.addParameter({name: 'Password', value: Password})};


// console.log("Page is loaded!");

ps.invoke()

.then(output => {

    console.log(output);
    var data;
    if (output) { 
    data = JSON.parse(output);}
    else {

        if (document.getElementById('alertbox').classList.contains('alert-danger')) {
            document.getElementById('alertbox').classList.remove('alert-danger');
            document.getElementById('alertbox').classList.add('alert-success');
        }
        document.getElementById('spinner').classList.toggle('d-none', true);
        document.getElementById('ErrorContainer').classList.remove('d-none');
        $('#AlertTitle').html("SEARCH COMPLETE");
        $('.alert-success .message').html("Your search for nothing was found. Change your search parameters and please try again.");
    }
  
    if (!$.isArray(data)) {data = [data]};
    
    console.log(data);

    document.getElementById('spinner').classList.add('d-none')
    document.getElementById('OutputContainer').classList.remove('d-none')


    // Create DataTable
    $('#output').DataTable({        
        data: data,        
        columns: [   
                           
           { title: "MessageId",
               data: "MessageId",
               className: "data-message",
               //render: function ( data, type, row, meta ) {
                //return type === 'display' && data.length > 40 ?
                  //'<span title="'+data+'">'+data.substr( 0, 38 )+'...</span>' : data;}

               render: function (mid) {
                   
               if (mid.length > 15 ){return ('<span class="tippy ' +mid +'">' + mid.substring(0,15) + '. . .</span>')}
               else {return mid};         
                              
               //className: "data-message-trns";
               }
            },
           { title: "Timestamp",
               data: "Timestamp", 
               render: function(d) {
               return moment(d).format("DD:MM:YYYY HH:mm:ss")}},
            { title: "Subject",
                data: "MessageSubject",
                render: function (mid) {                   
                    if (mid.length > 100 ){return ('<span class="tippy ' +mid +'">' + mid.substring(0,100) + '. . .</span>')}
                    else {return mid};
                }
            },
            { title: "EventId",
            data: "EventId",
            //className: "data-tippy-eventid"
            render: function (mid) {return ('<span class="event-tippy">' + mid + '</span>')}                    
            },

            { title: "ClientHostname",
                data: "ClientHostname" },  
            { title: "ServerHostname",
                data: "ServerHostname" },              
            { title: "Sender",
                data: "Sender" },
            { title: "Recipients",
                data: "Recipients" },
            { title: "Directionality",
                data: "Directionality" },
            { title: "Source",
                data: "Source",
                className: "data-tippy-source"                
            },            
            { title: "Size",
                data: "Size" }
        ],

        "columnDefs": [
            { className: "none", "targets": [ 4,5,6,7,8,9,10 ] }
          ],

        
        //serverSide: true,
        //processing: true,
        //deferLoading: 0,
        //searching: false,
        //info: false,
        //destroy: true,
        //scrollX:        true,
        //scrollCollapse: true,
        //autoWidth:         true,  
        //paging:         false, 

        destroy: true,
        responsive: {            
            details: {                
                renderer: function ( api, rowIdx, columns ) {
                    var data = $.map( columns, function ( col, i ) {
                        return col.hidden ?
                            '<tr data-dt-row="'+col.rowIndex+'" data-dt-column="'+col.columnIndex+'">'+
                                '<td>'+col.title+':'+'</td> '+
                                '<td>'+col.data+'</td>'+
                            '</tr>' :
                            '';
                    } ).join('');
 
                    return data ?
                        $('<table/>').append( data ) :
                        false;
                }
            }   
        },
        order: [[ 0, 'asc' ],[1, 'asc']],
        
        rowGroup: {
            startRender: function ( rows, group ) {
                if (group.length > 100 ){
                return group.substring(0,100) +'.. ('+rows.count()+')';}
                else {return group +'('+rows.count()+')'; }
            },
            endRender: null,
            dataSrc: "MessageId"
        }        
     
    }) //Create DataTable

    
   $('#output tbody').on('click', 'td.data-message', function () {
       console.log("loaded!");
      $('td.child > table > tr > td').each(function(){
        let reference = $(this).html();
        if (sourcetip[reference]){
            $(this).addClass('has-tip');
        }
        });
        tippy('.has-tip', {   
           content(reference) {
          if (sourcetip[reference.innerText]){
          const title = sourcetip[reference.innerText];        
           return title;
           }
           else {return ''};
          }
        });
        })
        
        
    $('#output tbody').on('mouseover', 'span.event-tippy', function () {
            console.log("lo!");
            tippy('.event-tippy', {   
                content(reference) {
                if (eventtip[reference.innerText]){
                const title = eventtip[reference.innerText];        
                return title;
                }
              }
            });

        });
    $('#output tbody').on('mouseover', 'span.tippy', function () {
            console.log("lo!");
            tippy('.tippy', { 
                content(reference) {
                    const title = reference.getAttribute('class').slice(5);
                    //reference.removeAttribute('title')
                    return title
                }
              })});
         
        

 


           
            
})
.catch(err => {
    console.log(err);

    document.getElementById('spinner').classList.toggle('d-none', true);
    if (!document.getElementById('OutputContainer').classList.contains('d-none')){
    document.getElementById('OutputContainer').classList.add('d-none');}
    if (document.getElementById('alertbox').classList.contains('alert-success')) {
        document.getElementById('alertbox').classList.remove('alert-success');
        document.getElementById('alertbox').classList.add('alert-danger');
    }
    document.getElementById('ErrorContainer').classList.remove('d-none');
    $('#AlertTitle').html("ERROR");
    $('.alert-danger .message').html(err);

  });
})


  


$('#returnbtn').click(() => {
    document.getElementById('OutputContainer').classList.add('d-none')
    document.getElementById('FormContainer').classList.remove('d-none')

})

$('#returnbtnerror').click(() => {
    $('.alert-danger .message').html();
    document.getElementById('ErrorContainer').classList.add('d-none')
    document.getElementById('FormContainer').classList.remove('d-none')

})
