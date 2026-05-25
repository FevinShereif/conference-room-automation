function updateClock() {

    const now = new Date();

    const options = {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    };

    document.getElementById("clock").innerHTML =
        now.toLocaleString('en-IN', options);
}

updateClock();

setInterval(updateClock, 1000);

setInterval(() => {

    location.reload();

}, 30000);


let pin = "";

function addPin(number) {

    if (pin.length < 6) {

        pin += number;

        document.getElementById("pin-input").value = pin;
    }
}

function clearPin() {

    pin = "";

    document.getElementById("pin-input").value = "";

    document.getElementById("pin-message").innerHTML = "";
}

function verifyPin() {

    const message =
        document.getElementById("pin-message");

    if (pin.length < 6) {

        message.innerHTML =
            "Enter 6-digit PIN";

        message.style.color = "#facc15";

        return;
    }

    message.innerHTML =
        "PIN verification successful";

    message.style.color = "#4ade80";

    setTimeout(() => {

        clearPin();

    }, 3000);
}