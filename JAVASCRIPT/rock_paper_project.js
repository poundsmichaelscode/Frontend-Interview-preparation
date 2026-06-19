
// Get compurt choice step one: generate random number between 0 and 2


function getComputerChoice() {
let randomNumber = math.random();

if (randomNumber < 0.34) {
    return "rock";}

    else if (randomNumber <= 0.67) {
        return "paper";}

        else {
            return "scissors";
        } 
}

// Get user choice step two: ask user for input and convert to lowercase

// console.log(getComputerChoice());
// console.log(getComputerChoice());
// console.log(getComputerChoice());
//get human choice step two: ask user for input and convert to lowercase

function getUserChoice(){

    return prompt("choose rock, paper, or scissors:").toLowerCase();
}


// console.log(getHumanChoice());


//get score
let humanScore = 0;

let computerScore= 0;

humanChoice = humanChoice.toLowerCase();

// play game step three: compare choices and determine winner


function playRound(humanChoice, computerChoice) {
    if (humanChoice === computerChoice) {
        return "It's a tie!";
    }

    else if (
  (humanChoice === "rock" && computerChoice === "scissors") ||
  (humanChoice === "paper" && computerChoice === "rock") ||
  (humanChoice === "scissors" && computerChoice === "paper")
) 

{
    humanScore++;
    return "You win! " + humanChoice + " beats " + computerChoice;
  } else {
    computerScore++;
    return "You lose! " + computerChoice + " beats " + humanChoice;
  }
 
}

console.log(`You win! ${humanChoice} beats ${computerChoice}`);