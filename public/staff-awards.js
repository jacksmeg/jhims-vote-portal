const nominationType = document.querySelector('select[name="nominationType"]');
const nominator = document.querySelector("[data-nominator]");

if (nominationType && nominator) {
  function updateNominator() {
    const required = nominationType.value === "another";
    nominator.hidden = !required;
    nominator.disabled = !required;
    for (const input of nominator.querySelectorAll("input")) input.required = required;
  }
  nominationType.addEventListener("change", updateNominator);
  updateNominator();
}
