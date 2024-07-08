$(document).ready(function() {
    const $extractForm = $('#extractForm');
    const $loadingMessage = $('#loadingMessage');
    const $successMessage = $('#successMessage');
    const $buttonContainer = $('.button-container');
    const $downloadLink = $('#downloadLink');
    const $newSearchButton = $('#newSearchButton');
    const $email = $('#email');

    $extractForm.submit(function(event) {
        event.preventDefault();
        $loadingMessage.show();
        $successMessage.hide();
        $buttonContainer.hide();

        $.ajax({
            url: '/extract',
            method: 'POST',
            data: $(this).serialize(),
            xhrFields: {
                responseType: 'blob'
            },
            success: function(response) {
                $loadingMessage.hide();
                $successMessage.show();
                $buttonContainer.show();

                const url = window.URL.createObjectURL(new Blob([response]));
                $downloadLink.attr('href', url);
            },
            error: function() {
                $loadingMessage.hide();
                alert('Ocorreu um erro ao processar a solicitação. Por favor, tente novamente mais tarde.');
            }
        });
    });

    $newSearchButton.click(function() {
        $email.val('');
        $successMessage.hide();
        $buttonContainer.hide();
    });
});