Office.onReady(() => {
    Office.actions.associate(
        "approveDocument",
        approveDocument
    );
});

function approveDocument(event) {
    const payload = {
        type: "approveDocument",
        source: "office-addin"
    };

    console.log("Posting approval message:", payload);

    if (window.parent && window.parent !== window) {
        window.parent.postMessage(payload, "*");
    } else {
        window.postMessage(payload, "*");
    }

    event.completed();
}