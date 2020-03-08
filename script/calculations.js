const observable = Rx.Observable.create(observer => {
    observer.next('hello');
    observer.next('world');
})

function print(val){
    let el = document.createElement('p');
    el.innerText = val;
    document.body.appendChild(el);
}

observable.subscribe(val => print(val));

const tableBody = document.getElementById("table-body");

var source = Rx.DOM.focusout(tableBody);

var subscription = source.subscribe(
    function (x) {
        alert('Next!');
    },
    function (err) {
        console.log('Error: ' + err);
    },
    function () {
        console.log('Completed');
    });

