function summarize() {
  Office.context.mailbox.item.body.getAsync("text", function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value;
      fetch("https://your-backend.com/summarize", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: emailBody })
      })
      .then(response => response.json())
      .then(data => {
        document.getElementById("summary").innerText = data.summary;
      });
    }
  });
}
