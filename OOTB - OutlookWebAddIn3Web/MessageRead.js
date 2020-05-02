(function () {
  "use strict";

  var customProps;
  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      var element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
        loadProps();

        var btn1 = document.querySelector('#btn1');
        btn1.addEventListener("click", newThing);
        btn1.textContent = "Ready to popup";

        var btnSave = document.querySelector('#btnSave');
        btnSave.addEventListener("click", saveRoaming);
        btnSave.textContent = "Ready to Save Roaming";

        var btnSC = document.querySelector('#btnSC');
        btnSC.addEventListener("click", setCustom);
        btnSC.textContent = "Ready to Save Custom";

        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, SelectedItemChanged);

        var value = Office.context.roamingSettings.get('myKey');
        $('#dvmyKey').html(value);
        loadCats();

        loadCustomProps();
        
    });
    };

    function setCustom() {
        customProps.set('a', 'b');
        saveCustom();
    }

    function saveCustom() {
        customProps.saveAsync(function (result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                //console.error(`saveAsync failed with message ${result.error.message}`);
            } else {
                //console.log(`Custom properties saved with status: ${result.status}`);
            }
        });
    }

    function loadCustomProps() {
        Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                $('#custom').append("Loaded following custom properties:");
                customProps = result.value;
                var dataKey = Object.keys(customProps)[0];
                var data = customProps[dataKey];
                for (var propertyName in data) {
                    var propertyValue = data[propertyName];
                    $('#custom').append(propertyName + " : " + propertyValue);
                }
            }
            else {
                $('#custom').append('loadCustomPropertiesAsync failed with message ${result.error.message}');
            }
        });

    }

    function loadCats() {

        $('#cats').html("Categories");
        //$('#cats').append(Office.context.mailbox.masterCategories);

        Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                 $('#cats').append(asyncResult.error.message);
            } else {
                var masterCategories = asyncResult.value;
                masterCategories.forEach(function (item) {

                    $('#cats').append(item);

                });
            }
        });      

     }

    function SelectedItemChanged() {
        loadProps();
        loadCats();
        var value = Office.context.roamingSettings.get('myKey');
        $('#dvmyKey').html(value);
        $('#custom').html('');
        loadCustomProps();
        
    }

    function saveRoaming() {
        Office.context.roamingSettings.set('myKey', 'I am set: ' + Date());
        Office.context.roamingSettings.saveAsync();
        var value = Office.context.roamingSettings.get('myKey');
        $('#dvmyKey').html(value);
    }

    function newThing() {
        Office.context.ui.displayDialogAsync("https://www.bliss-systems.co.uk",
            { height: 30, width: 20, displayInIframe: false ,promptBeforeOpen:false},
            function (asyncResult) {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            }
        );
    }

    function processMessage(arg) {
        var messageFromDialog = JSON.parse(arg.message);
        showUserName(messageFromDialog.name);
    }
  // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>
  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";

      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }

      return returnString;
    }

    return "None";
  }

  // Load properties from the Item base object, then load the
  // message-specific properties.
  function loadProps() {
    var item = Office.context.mailbox.item;

    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    $('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
      $('#to').html(buildEmailAddressesString(item.to));

      $('#MailBoxAddress').html(Office.context.mailbox.userProfile.emailAddress);
      //$('#MailBoxID').html(Office.context.mailbox.);
      
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();