$(document).ready(function() {
    const $searchButton = $('#searchButton');
    const $extractForm = $('#extractForm');
    const $loadingMessage = $('#loadingMessage');
    const $successMessage = $('#successMessage');
    const $errorMessage = $('#errorMessage');
    const $downloadButton = $('#download-btn');
    const $newSearchButton = $('#newSearchButton');
    const $email = $('#email');

    $extractForm.submit(function(event) {
        event.preventDefault();
        const email = $email.val().trim();

        if (!isValidEmail(email)) {
            alert('Por favor, insira um e-mail válido.');
            return;
        }

        $loadingMessage.show();
        $successMessage.hide();
        $errorMessage.hide();
        $downloadButton.hide();
        $newSearchButton.hide();

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
                $downloadButton.show();
                $newSearchButton.show();
                $searchButton.hide();
                
                const url = window.URL.createObjectURL(new Blob([response]));
                const $downloadLink = $('#downloadLink');
                $downloadLink.attr('href', url);
            },
            error: function(xhr) {
                $loadingMessage.hide();
                if (xhr.status === 404) {
                    $errorMessage.text('Nenhum dado encontrado para exportar.').show();
                } else {
                    $errorMessage.text('Ocorreu um erro ao processar a solicitação. Por favor, tente novamente mais tarde.').show();
                }
                $searchButton.hide();
                $newSearchButton.show();
            }
        });
    });

    $newSearchButton.click(function() {
        $email.val('');
        $successMessage.hide();
        $errorMessage.hide();
        $downloadButton.hide();
        $newSearchButton.hide();
        $searchButton.show();
    });

    function isValidEmail(email) {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return emailRegex.test(email);
    }
});