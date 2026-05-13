const pet = document.querySelector("#pet");
const speechBubble = document.querySelector("#speechBubble");
const moodLabel = document.querySelector("#moodLabel");
const energyMeter = document.querySelector("#energyMeter");
const feedButton = document.querySelector("#feedButton");
const playButton = document.querySelector("#playButton");
const napButton = document.querySelector("#napButton");
const resetButton = document.querySelector("#resetButton");

const moods = {
  curious: {
    label: "好奇",
    lines: ["这里闻起来像新点子！", "你在忙什么呀？带我一个！", "双击我可以换表情哦。"],
  },
  happy: {
    label: "开心",
    lines: ["星星好甜！谢谢你～", "今天的幸福值加满啦！", "嘿嘿，我会发光吗？"],
  },
  playful: {
    label: "调皮",
    lines: ["来追我呀！", "我刚刚是不是跳得很高？", "再玩五分钟好不好？"],
  },
  sleepy: {
    label: "困困",
    lines: ["晚安，云朵要充电了……", "呼噜呼噜。", "帮我把梦收好。"],
  },
};

const state = {
  mood: "curious",
  energy: 72,
  dragging: false,
  pointerOffsetX: 0,
  pointerOffsetY: 0,
};

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function randomItem(items) {
  return items[Math.floor(Math.random() * items.length)];
}

function setPetPosition(x, y) {
  const safeX = clamp(x, 8, window.innerWidth - pet.offsetWidth - 8);
  const safeY = clamp(y, 8, window.innerHeight - pet.offsetHeight - 8);
  pet.style.setProperty("--pet-x", `${safeX}px`);
  pet.style.setProperty("--pet-y", `${safeY}px`);
}

function speak(message) {
  speechBubble.textContent = message;
}

function burstSparkles(count = 7) {
  const rect = pet.getBoundingClientRect();

  for (let index = 0; index < count; index += 1) {
    const sparkle = document.createElement("span");
    sparkle.className = "sparkle";
    sparkle.textContent = index % 2 === 0 ? "✦" : "★";
    sparkle.style.left = `${rect.left + rect.width * (0.2 + Math.random() * 0.6)}px`;
    sparkle.style.top = `${rect.top + rect.height * (0.2 + Math.random() * 0.5)}px`;
    sparkle.style.animationDelay = `${index * 45}ms`;
    document.body.append(sparkle);
    sparkle.addEventListener("animationend", () => sparkle.remove(), { once: true });
  }
}

function setMood(nextMood, energyDelta = 0, customLine) {
  const mood = moods[nextMood] ?? moods.curious;
  state.mood = nextMood;
  state.energy = clamp(state.energy + energyDelta, 0, 100);

  pet.classList.remove(...Object.keys(moods));
  pet.classList.add(nextMood, "pop");
  moodLabel.textContent = mood.label;
  energyMeter.value = state.energy;
  energyMeter.textContent = String(state.energy);
  speak(customLine ?? randomItem(mood.lines));

  pet.addEventListener("animationend", () => pet.classList.remove("pop"), { once: true });
}

function resetPet() {
  setPetPosition(window.innerWidth - pet.offsetWidth - 48, window.innerHeight - pet.offsetHeight - 48);
  setMood("curious", 0, "我回到舒适的小角落啦。");
}

function feedPet() {
  setMood("happy", 16);
  burstSparkles(10);
}

function playWithPet() {
  const x = 24 + Math.random() * Math.max(24, window.innerWidth - pet.offsetWidth - 72);
  const y = 80 + Math.random() * Math.max(24, window.innerHeight - pet.offsetHeight - 140);
  setPetPosition(x, y);
  setMood("playful", -9);
  burstSparkles(5);
}

function napPet() {
  setMood("sleepy", 22);
}

function rotateMood() {
  const moodOrder = Object.keys(moods);
  const currentIndex = moodOrder.indexOf(state.mood);
  const nextMood = moodOrder[(currentIndex + 1) % moodOrder.length];
  setMood(nextMood, nextMood === "playful" ? -4 : 3);
}

pet.addEventListener("pointerdown", (event) => {
  state.dragging = true;
  pet.classList.add("dragging");
  pet.setPointerCapture(event.pointerId);
  const rect = pet.getBoundingClientRect();
  state.pointerOffsetX = event.clientX - rect.left;
  state.pointerOffsetY = event.clientY - rect.top;
  speak("带我去新的地方看看！");
});

pet.addEventListener("pointermove", (event) => {
  if (!state.dragging) {
    return;
  }

  setPetPosition(event.clientX - state.pointerOffsetX, event.clientY - state.pointerOffsetY);
});

pet.addEventListener("pointerup", (event) => {
  state.dragging = false;
  pet.classList.remove("dragging");
  pet.releasePointerCapture(event.pointerId);
  setMood("curious", -1, "这里的风景不错！");
});

pet.addEventListener("dblclick", rotateMood);
pet.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    rotateMood();
  }
});

feedButton.addEventListener("click", feedPet);
playButton.addEventListener("click", playWithPet);
napButton.addEventListener("click", napPet);
resetButton.addEventListener("click", resetPet);
window.addEventListener("resize", resetPet);

setMood("curious", 0, "嗨！今天也要元气满满～");
resetPet();
