//Function in Javascript


// function num (a, b){

// console.log(a + b )
// }



// num(10, 20)



// var colors = ['red', 'green', 'blue', 'yellow', 'orange']

// arrayItems(colors)



// function arrayItems(array) {
//     for (let i = 0; i < array.length; i++) {
//         console.log(array[i])
//     }
// }   

// arrayItems(colors)




function letterFinder (word, match) {

    for (let i = 0; i < word.length; i++) {
      

        if (word[i] === match) {
            console.log(`Found the ${match} at ${i}`)
        } 
        
        else {
            console.log(`No match found at ${i}`)
        }
      
    }

}


//   word[i] === match ? console.log(`Found the ${match} at ${i}`) : console.log(`No match found at ${i}`)