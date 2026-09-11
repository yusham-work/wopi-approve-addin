Office.onReady(() => {
    Office.actions.associate(
        "approveDocument",
        approveDocument
    );
});


async function approveDocument(event) {
    try {
        console.log("Approve button clicked");

        const response = await fetch(
            "https://myapp.example.com/api/documents/approve",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    source: "office-addin"
                })
            }
        );

        if (!response.ok) {
            throw new Error(
                `HTTP ${response.status}`
            );
        }

        console.log("Document approved");

    } catch (error) {
        console.error(
            "Approval failed:",
            error
        );

    } finally {
        event.completed();
    }
}