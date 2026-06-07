//Boolean (BOO-lee-uhn) is another data type in JavaScript
.// JavaScript boolean values can be either true or false
// They are useful for determining whether or not blocks of code should be executed. They are also the default result for many evaluations.

let likesJavaScript = true;
let likesMath = false;

const numToCheck = 10;
console.log(numToCheck === 10) // => true


//====================Comparison Operatore =====================

Comparison Operators
In programming, when working with data, we often need to compare different values. To do this, we use a series of operators called comparison operators. Check out this list of the most common comparison operators:

< - Less than
> - Greater than
<= - Less than or equal to
>= - Greater than or equal to
== - Is loosely equal to
=== - Is strictly equal to
!= - Is not loosely equal to
!== - Is not strictly equal to
The first four are probably pretty familiar to you from primary school math class, but things start to get a little tricky when we talk about equality in JavaScript.

First, we already know that a single equals sign (=) is used to assign value to a variable, so we can't use that to compare to values unfortunately.

So let's start with loose equality (==). This operator is used to compare if 2 values have the same value, even if they aren't necessarily the same type

// 1.
const small = 2;
const large = 5342;
// ADD CODE BELOW (isSmaller)

let isSmaller = small < large;

// 2.
const num = 45;
const string = "45";
// ADD CODE BELOW (isLooselyEqual and isStrictlyEqual)

let  isLooselyEqual = num == string;

// 3.
const isTrue = true;
const isFalse = false;
// ADD CODE BELOW (isTrueNotFalse)
let isTrueNotFalse =  isTrue  !== isFalse
let isStrictlyEqual = num === string
// Uncomment these to check your work! 
console.log('Is 2 < 5342?');
console.log(isSmaller);
console.log('Is 45 loosely equal to "45"?'); 
console.log(isLooselyEqual);
console.log('Is 45 strictly equal to "45"?');
console.log(isStrictlyEqual);
console.log('Is true not equal to false?');
console.log(isTrueNotFalse);.
