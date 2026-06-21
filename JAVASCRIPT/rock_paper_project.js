// Function to get computer's choice


function getComputerChoice() {
  let randomNumber = Math.random();

  if (randomNumber < 0.33) {
    return "rock";
  } else if (randomNumber < 0.66) {
    return "paper";
  } else {
    return "scissors";
  }
}


//get human choice from prompt

function getHumanChoice() {
   prompt("Enter rock, paper, or scissors:");
}


// Function to play the game
function playGame() {
  let humanScore = 0;
  let computerScore = 0;

  // Function to play a single round

  function playRound(humanChoice, computerChoice) {
    humanChoice = humanChoice.toLowerCase();

    if (humanChoice === computerChoice) {
      console.log("It's a tie!");
    } else if (
      (humanChoice === "rock" && computerChoice === "scissors") ||
      (humanChoice === "paper" && computerChoice === "rock") ||
      (humanChoice === "scissors" && computerChoice === "paper")
    )
    {
      humanScore++;
      console.log(`You win! ${humanChoice} beats ${computerChoice}`);
    } else {
      computerScore++;
      console.log(`You lose! ${computerChoice} beats ${humanChoice}`);
    }
  }

  playRound(getHumanChoice(), getComputerChoice());
  playRound(getHumanChoice(), getComputerChoice());
  playRound(getHumanChoice(), getComputerChoice());
  playRound(getHumanChoice(), getComputerChoice());
  playRound(getHumanChoice(), getComputerChoice());

  console.log(`Human Score: ${humanScore}`);
  console.log(`Computer Score: ${computerScore}`);

  if (humanScore > computerScore) {
    console.log("You won the game!");
  } else if (computerScore > humanScore) {
    console.log("Computer won the game!");
  } else {
    console.log("The game ended in a tie!");
  }
}

playGame();