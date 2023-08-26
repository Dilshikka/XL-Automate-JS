// const nextButton = document.getElementById('nextButton');
const container = document.querySelector('.page_one_wrapper');
const newContainer = document.querySelector('.page_two_wrapper');

nextButton.addEventListener('click', () => {
    // Delay to match the transition duration in CSS
});

function nextPage() {

    container.innerHTML = "<h1> new page</h1>"

    // container.classList.add('showNewContainer');
    // setTimeout(() => {
    //     container.style.display = 'none';
    //     newContainer.classList.add('show');
    // }, 500);
}
