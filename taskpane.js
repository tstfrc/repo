Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const item = Office.context.mailbox.item;

    // Esempio: ottieni oggetto e data
    Promise.all([
      new Promise((resolve) => item.subject.getAsync(resolve)),
      new Promise((resolve) => item.start.getAsync(resolve))
    ]).then(([subject, start]) => {
      const data = {
        subject: subject.value,
        start: start.value
      };

      // Esegui chiamata API
      fetch("https://eu.celoxis.com/psa/api/v2/projects", {
        method: "POST",
        headers: {
          "Authorization": "Bearer JS7qrYJpbJGp0n0WZ8eu11QES5FLBM61bIOpOstW",
          "Content-Type": "application/json"
        },
        body: JSON.stringify(data)
      })
      .then(res => res.json())
      .then(json => {
        document.getElementById("output").textContent = JSON.stringify(json, null, 2);
      })
      .catch(err => {
        document.getElementById("output").textContent = "Errore API: " + err;
      });
    });
  }
});