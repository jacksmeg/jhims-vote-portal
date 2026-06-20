(function () {
  const root = document.querySelector("[data-demo-root]");
  if (!root) return;

  const screens = Array.from(root.querySelectorAll("[data-demo-screen]"));
  const roadmapSteps = Array.from(root.querySelectorAll("[data-roadmap-step]"));
  const mobileSteps = Array.from(root.querySelectorAll("[data-mobile-step]"));
  const positions = Array.from(root.querySelectorAll("[data-demo-position]"));
  const selections = positions.map(() => null);
  const stageOrder = ["login", "otp", "ballot", "review", "complete"];
  let currentPositionIndex = 0;

  function showScreen(name) {
    screens.forEach((screen) => screen.classList.toggle("is-active", screen.dataset.demoScreen === name));
    const activeIndex = stageOrder.indexOf(name);

    roadmapSteps.forEach((step, index) => {
      step.classList.toggle("is-current", index === activeIndex);
      step.classList.toggle("is-complete", index < activeIndex);
    });
    mobileSteps.forEach((step, index) => {
      step.classList.toggle("is-current", index === activeIndex);
      step.classList.toggle("is-complete", index < activeIndex);
    });

    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function showPosition(index) {
    currentPositionIndex = Math.min(Math.max(index, 0), positions.length - 1);
    positions.forEach((position, positionIndex) => {
      position.classList.toggle("is-active", positionIndex === currentPositionIndex);
    });

    const position = positions[currentPositionIndex];
    root.querySelector("[data-demo-position-name]").textContent = position.dataset.positionName;
    root.querySelector("[data-demo-ballot-step]").textContent = `Position ${currentPositionIndex + 1} of ${positions.length}`;
    root.querySelector("[data-demo-progress]").style.width = `${((currentPositionIndex + 1) / positions.length) * 100}%`;
    root.querySelector("[data-demo-completed]").textContent = selections.filter((selection) => selection !== null).length;
    root.querySelector("[data-demo-ballot-back]").disabled = currentPositionIndex === 0;
    root.querySelector("[data-demo-next]").textContent = currentPositionIndex === positions.length - 1 ? "Review Ballot" : "Next Position";
    root.querySelector("[data-demo-skip]").textContent = currentPositionIndex === positions.length - 1 ? "Skip and Review" : "Skip Position";
    position.querySelector("[data-demo-ballot-error]").textContent = "";
  }

  function saveChoice(choice) {
    selections[currentPositionIndex] = choice;
    if (currentPositionIndex === positions.length - 1) {
      renderReview();
      showScreen("review");
      return;
    }
    showPosition(currentPositionIndex + 1);
  }

  function renderReview() {
    const list = root.querySelector("[data-demo-review-list]");
    list.replaceChildren();

    positions.forEach((position, index) => {
      const row = document.createElement("article");
      const copy = document.createElement("div");
      const positionName = document.createElement("strong");
      const label = document.createElement("span");
      const choice = document.createElement("strong");

      positionName.textContent = position.dataset.positionName;
      label.textContent = selections[index]?.skipped ? "Position skipped" : "Selected candidate";
      choice.textContent = selections[index]?.skipped ? "No selection" : selections[index]?.candidate || "Not completed";
      choice.className = selections[index]?.skipped ? "is-skipped" : "";
      copy.append(positionName, label);
      row.append(copy, choice);
      list.append(row);
    });
  }

  root.querySelector("[data-demo-login-form]").addEventListener("submit", (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const staffId = form.elements.staffId.value.trim();
    const phoneNumber = form.elements.phoneNumber.value.trim();
    const error = root.querySelector("[data-demo-login-error]");

    if (!staffId || !phoneNumber) {
      error.textContent = "Enter both sample details to continue the demonstration.";
      return;
    }
    error.textContent = "";
    showScreen("otp");
  });

  root.querySelector("[data-demo-otp-form]").addEventListener("submit", (event) => {
    event.preventDefault();
    const code = event.currentTarget.elements.otp.value.trim();
    const error = root.querySelector("[data-demo-otp-error]");
    if (code !== "2026") {
      error.textContent = "For this training demo, enter the code 2026.";
      return;
    }
    error.textContent = "";
    showScreen("ballot");
    showPosition(0);
  });

  root.querySelectorAll("[data-demo-back]").forEach((button) => {
    button.addEventListener("click", () => showScreen(button.dataset.demoBack));
  });

  root.querySelector("[data-demo-next]").addEventListener("click", () => {
    const activePosition = positions[currentPositionIndex];
    const selected = activePosition.querySelector("input[type='radio']:checked");
    if (!selected) {
      activePosition.querySelector("[data-demo-ballot-error]").textContent = "Select a candidate or use Skip Position.";
      return;
    }
    saveChoice({ candidate: selected.value, skipped: false });
  });

  root.querySelector("[data-demo-skip]").addEventListener("click", () => {
    positions[currentPositionIndex].querySelectorAll("input[type='radio']").forEach((input) => {
      input.checked = false;
    });
    saveChoice({ candidate: "", skipped: true });
  });

  root.querySelector("[data-demo-ballot-back]").addEventListener("click", () => {
    if (currentPositionIndex > 0) showPosition(currentPositionIndex - 1);
  });

  root.querySelector("[data-demo-review-back]").addEventListener("click", () => {
    showScreen("ballot");
    showPosition(positions.length - 1);
  });

  const dialog = root.querySelector("[data-demo-dialog]");
  root.querySelector("[data-demo-open-confirm]").addEventListener("click", () => dialog.showModal());
  root.querySelector("[data-demo-cancel-confirm]").addEventListener("click", () => dialog.close());
  root.querySelector("[data-demo-confirm-submit]").addEventListener("click", () => {
    dialog.close();
    showScreen("complete");
  });

  root.querySelectorAll("[data-demo-reset]").forEach((button) => {
    button.addEventListener("click", () => {
      selections.fill(null);
      currentPositionIndex = 0;
      root.querySelectorAll("form").forEach((form) => form.reset());
      root.querySelectorAll("[data-demo-ballot-error], [data-demo-login-error], [data-demo-otp-error]").forEach((error) => {
        error.textContent = "";
      });
      showPosition(0);
      showScreen("login");
    });
  });

  root.querySelectorAll(".vote-demo__candidate input").forEach((input) => {
    input.addEventListener("change", () => {
      const position = input.closest("[data-demo-position]");
      position.querySelector("[data-demo-ballot-error]").textContent = "";
    });
  });
})();
