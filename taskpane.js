function summarize() {
  const summaryBox = document.getElementById("summaryBox");
  const copyButton = document.getElementById("copyButton");

  summaryBox.innerText = "Summarizing...";
  summaryBox.classList.add("loading");
  copyButton.style.display = "none";

  Office.context.mailbox.item.body.getAsync("text", function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value;

      fetch("https://summarize-backend.onrender.com/summarize", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: emailBody })
      })
      .then(response => response.json())
      .then(data => {
        summaryBox.innerText = data.summary || "No summary returned.";
        summaryBox.classList.remove("loading");
        copyButton.style.display = "inline-block";
      })
      .catch(err => {
        summaryBox.innerText = "Error summarizing email.";
        summaryBox.classList.remove("loading");
        console.error(err);
      });
    }
  });
}

function copySummary() {
  const summaryText = document.getElementById("summaryBox").innerText;
  navigator.clipboard.writeText(summaryText).then(() => {
    const copyBtn = document.getElementById("copyButton");
    copyBtn.innerText = "Copied!";
    copyBtn.classList.add("copied");

    setTimeout(() => {
      copyBtn.innerText = "Copy Summary";
      copyBtn.classList.remove("copied");
    }, 1500);
  });
}
