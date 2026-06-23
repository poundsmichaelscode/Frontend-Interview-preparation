
for (let i =0; i <=10; i++)

{

console.log(i);
}





for (let i =3; i >=  0; i--)



{
console.log(i);

}



console.log ("count 1-5")


for (let count = 1; count <= 5; count++) {
  console.log(count);



}



console.log("divide by 2")


for (i =0; i<= 20 ;  i+=2)

{
console.log(i)
};


console.log ("by 3")


for (i =0; i <=18 ; i +=3)


{console.log (i)};





console.log ("count down")


for (let i =3; i>=0; i--)

{ console.log(i)

}



console.log ("print stars");


let stars ="";


for (let star= 0;   star<=8;  star++)

{ stars += "***"


console.log (stars);

}










for (i = 5; i >=0 ; i--){

console.log(i)};



// increment by 2

 


for (let i = 2; i <= 10; i+=2){

console.log(i)

}

console.log("loop by two")


for (let i =10; i >0 ; i-=2)
{console.log(i)}

console.log ("hello world!!")

console.log ("GOT")



const firstNames = ["Jon", "Arya", "Jamie"];
const lastNames = ["Snow", "Stark", "Lannister"];
const places = ["The Wall", "Winterfell", "Kings Landing"];


const bios = [];

// Loop through your arrays and store the following strings in the bios array:
// 'My name is Jon Snow and I am from The Wall'
// 'My name is Arya Stark and I am from Winterfell'
// 'My name is Jamie Lannister and I am from Kings Landing'

// ADD CODE HERE


for (let i =0; i < firstNames.length; i++)
{
  
bios.push(`My name is ${firstNames[i]} ${lastNames[i]} and I am from the ${places[i]}`);
  
}


console.log(bios)



console.log("While loop")

let counter = 1;

while (counter <=5 ) {

    console.log(counter);
     counter++
}

console.log('Counting completed!')
