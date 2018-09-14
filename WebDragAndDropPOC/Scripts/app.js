function DragDropViewModel() {
    var self = this;

    self.dragover = function (e) {
        console.log('dragOver');       
        self.hover(true);
        e.stopPropagation();
        e.preventDefault();
    }

    self.drop = function (e, data) {
        console.log('drop');
        self.hover(false);
        e.stopPropagation();
        e.preventDefault();
    }

    self.dragenter = function (e) {
        console.log('dragEnter');
    }

    self.dragleave = function (e) {
        console.log('end');
        self.hover(false);
    }

    self.hover = new ko.observable(false);
}

function ResponseViewModel() {
    var self = this;
    self.fileDetails = ko.observableArray([]);
}

function DetailsViewModel() {
    var self = this;

    self.filename = ko.observable('');
    self.subject = ko.observable('');
    self.sender = ko.observable('');
    self.recipients = ko.observable('');
    self.attachments = ko.observableArray([]);
}


$(function () {

    //prevent drag&drop effects on rest of page
    window.addEventListener("dragover", function (e) {
        e = e || event;
        e.preventDefault();
    }, false);
    window.addEventListener("drop", function (e) {
        e = e || event;
        e.preventDefault();
    }, false);

    //set up KO view models
    var dragDropViewModel = new DragDropViewModel();
    var responseViewModel = new ResponseViewModel();
    ko.applyBindings(dragDropViewModel, document.getElementById('dropzone'));
    ko.applyBindings(responseViewModel, document.getElementById('fileDetails'));

    //configure upload plugin
    $('#fileupload').fileupload({
        url: '/api/msgfile',
        dataType: 'json',
        dropZone: $('#dropzone'),
        singleFileUploads: false,
        add: function (e, data) {
            data.submit();
        },
        done: function (e, data) {
            var messageFiles = data._response.result;
            var detailsArray = new Array();

            $.each(messageFiles, function (index, value) {

                var detailsViewModel = new DetailsViewModel();

                detailsViewModel.filename(value.FileName);
                detailsViewModel.subject(value.Subject);
                detailsViewModel.sender(value.Sender);
                detailsViewModel.recipients(value.RecipientsTo);
                detailsViewModel.attachments(value.Attachments);
                detailsArray.push(detailsViewModel);
            });

            responseViewModel.fileDetails(detailsArray);
        }
    });
});