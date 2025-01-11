document.addEventListener('DOMContentLoaded', () => {
    // Function to handle conversion requests
    async function selectConversion(type) {
        // Create a file input dynamically
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = getFileAcceptType(type);

        // Trigger file selection
        fileInput.onchange = async () => {
            const file = fileInput.files[0];
            if (!file) {
                alert('No file selected. Please try again.');
                return;
            }

            // Prepare form data
            const formData = new FormData();
            formData.append('file', file);
            formData.append('type', type);

            try {
                // Send the file to the backend for processing
                const response = await fetch('/upload', {
                    method: 'POST',
                    body: formData,
                });

                // Handle the server response
                const result = await response.json();
                if (response.ok) {
                    alert(result.message);
                    if (result.path) {
                        window.open(result.path, '_blank'); // Open the converted PDF
                    }
                } else {
                    alert(result.error || 'Something went wrong!');
                }
            } catch (error) {
                console.error('Error:', error);
                alert('An error occurred while processing your file. Please try again.');
            }
        };

        // Open the file dialog
        fileInput.click();
    }

    // Helper function to return the correct file type for input
    function getFileAcceptType(type) {
        switch (type) {
            case 'word':
                return '.doc, .docx';
            case 'excel':
                return '.xls, .xlsx';
            case 'ppt':
                return '.ppt, .pptx';
            case 'jpg':
                return '.jpg, .jpeg, .png';
            default:
                return '*/*'; // Fallback for any file type
        }
    }

    // Attach event listeners to buttons
    document.querySelectorAll('.btn').forEach(button => {
        button.addEventListener('click', () => {
            const type = button.getAttribute('onclick').match(/'(\w+)'/)[1];
            selectConversion(type);
        });
    });
});
