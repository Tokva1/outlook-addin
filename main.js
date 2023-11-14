// Replace 'YOUR_GPT4_API_KEY' with your actual GPT-4 API key
const GPT4_API_KEY = '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';

document.getElementById('summarize-btn').addEventListener('click', function () {
    Office.context.mailbox.item.body.getAsync('text', (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = result.value;
            summarizeEmail(emailBody);
        } else {
            // Handle the error
            console.error(result.error);
        }
    });
});

function summarizeEmail(emailBody) {
    const requestOptions = {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${GPT4_API_KEY}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            prompt: emailBody,
            max_tokens: 200
        })
    };

    fetch('https://api.openai.com/v1/engines/davinci-codex/completions', requestOptions)
        .then(response => response.json())
        .then(data => {
            document.getElementById('summary').innerText = data.choices[0].text;
            suggestReplies(emailBody); // Now using the original email body for suggestions
        })
        .catch(error => console.error('Error:', error));
}

function suggestReplies(emailBody) {
    // Adjust this function to call GPT-4 for generating reply suggestions
    const requestOptions = {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${GPT4_API_KEY}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            prompt: "Suggest three reply options for this email: \"" + emailBody + "\"",
            max_tokens: 150 // Adjust as needed
        })
    };

    fetch('https://api.openai.com/v1/engines/davinci-codex/completions', requestOptions)
        .then(response => response.json())
        .then(data => {
            const replies = data.choices[0].text.trim().split('\n'); // Assuming each reply is on a new line
            displayReplies(replies);
        })
        .catch(error => console.error('Error:', error));
}

function displayReplies(replies) {
    const repliesElement = document.getElementById('suggested-replies');
    repliesElement.innerHTML = '<p>Suggested Replies:</p><ul>' + replies.map(reply => `<li>${reply}</li>`).join('') + '</ul>';
}
