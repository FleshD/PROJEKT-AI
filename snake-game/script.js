const canvas = document.querySelector('#game');
const context = canvas.getContext('2d');
const scoreElement = document.querySelector('#score');
const bestScoreElement = document.querySelector('#bestScore');
const startButton = document.querySelector('#startButton');
const messageElement = document.querySelector('#message');
const touchButtons = document.querySelectorAll('[data-direction]');

const tileCount = 24;
const tileSize = canvas.width / tileCount;
const startPosition = { x: 12, y: 12 };
const directions = {
  ArrowUp: { x: 0, y: -1 },
  KeyW: { x: 0, y: -1 },
  ArrowDown: { x: 0, y: 1 },
  KeyS: { x: 0, y: 1 },
  ArrowLeft: { x: -1, y: 0 },
  KeyA: { x: -1, y: 0 },
  ArrowRight: { x: 1, y: 0 },
  KeyD: { x: 1, y: 0 },
};

let snake;
let apple;
let velocity;
let nextVelocity;
let score;
let bestScore = Number(localStorage.getItem('snakeBestScore')) || 0;
let gameLoop;
let isRunning = false;

bestScoreElement.textContent = bestScore;
resetGame();
draw();

startButton.addEventListener('click', startGame);
document.addEventListener('keydown', handleKeyDown);
touchButtons.forEach((button) => {
  button.addEventListener('click', () => setDirection(button.dataset.direction));
});

function startGame() {
  resetGame();
  isRunning = true;
  messageElement.textContent = 'Hra běží! Sbírej červená jablka.';
  clearInterval(gameLoop);
  gameLoop = setInterval(updateGame, 115);
}

function resetGame() {
  snake = [
    { ...startPosition },
    { x: startPosition.x - 1, y: startPosition.y },
    { x: startPosition.x - 2, y: startPosition.y },
  ];
  velocity = { x: 1, y: 0 };
  nextVelocity = { x: 1, y: 0 };
  score = 0;
  scoreElement.textContent = score;
  apple = placeApple();
}

function updateGame() {
  velocity = nextVelocity;
  const head = {
    x: snake[0].x + velocity.x,
    y: snake[0].y + velocity.y,
  };

  if (hasCollision(head)) {
    endGame();
    return;
  }

  snake.unshift(head);

  if (head.x === apple.x && head.y === apple.y) {
    score += 10;
    scoreElement.textContent = score;
    apple = placeApple();
  } else {
    snake.pop();
  }

  draw();
}

function hasCollision(head) {
  const hitWall = head.x < 0 || head.y < 0 || head.x >= tileCount || head.y >= tileCount;
  const hitSnake = snake.some((part) => part.x === head.x && part.y === head.y);
  return hitWall || hitSnake;
}

function endGame() {
  clearInterval(gameLoop);
  isRunning = false;
  bestScore = Math.max(bestScore, score);
  localStorage.setItem('snakeBestScore', bestScore);
  bestScoreElement.textContent = bestScore;
  messageElement.textContent = `Konec hry! Tvoje skóre: ${score}. Zkus to znovu.`;
  draw();
}

function placeApple() {
  let position;

  do {
    position = {
      x: Math.floor(Math.random() * tileCount),
      y: Math.floor(Math.random() * tileCount),
    };
  } while (snake.some((part) => part.x === position.x && part.y === position.y));

  return position;
}

function draw() {
  context.fillStyle = '#07130d';
  context.fillRect(0, 0, canvas.width, canvas.height);
  drawGrid();
  drawTile(apple, '#ef4444', '#fecaca');

  snake.forEach((part, index) => {
    drawTile(part, index === 0 ? '#86efac' : '#22c55e', index === 0 ? '#dcfce7' : '#bbf7d0');
  });

  if (!isRunning) {
    context.fillStyle = 'rgba(0, 0, 0, 0.36)';
    context.fillRect(0, 0, canvas.width, canvas.height);
  }
}

function drawGrid() {
  context.strokeStyle = 'rgba(255, 255, 255, 0.04)';
  context.lineWidth = 1;

  for (let line = 0; line <= tileCount; line += 1) {
    const position = line * tileSize;
    context.beginPath();
    context.moveTo(position, 0);
    context.lineTo(position, canvas.height);
    context.moveTo(0, position);
    context.lineTo(canvas.width, position);
    context.stroke();
  }
}

function drawTile(tile, fill, stroke) {
  const padding = 2;
  context.fillStyle = fill;
  context.strokeStyle = stroke;
  context.lineWidth = 2;
  context.beginPath();
  context.roundRect(
    tile.x * tileSize + padding,
    tile.y * tileSize + padding,
    tileSize - padding * 2,
    tileSize - padding * 2,
    6,
  );
  context.fill();
  context.stroke();
}

function handleKeyDown(event) {
  if (directions[event.code]) {
    event.preventDefault();
    setNextVelocity(directions[event.code]);
  }
}

function setDirection(direction) {
  const mapping = {
    up: directions.ArrowUp,
    down: directions.ArrowDown,
    left: directions.ArrowLeft,
    right: directions.ArrowRight,
  };
  setNextVelocity(mapping[direction]);
}

function setNextVelocity(direction) {
  if (!direction) {
    return;
  }

  const isOppositeDirection = direction.x + velocity.x === 0 && direction.y + velocity.y === 0;

  if (!isOppositeDirection) {
    nextVelocity = direction;
  }
}
