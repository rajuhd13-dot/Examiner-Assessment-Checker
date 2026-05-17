const url = "https://script.google.com/macros/s/AKfycbxun3tEoQdPq4vrBT6Xk-paLYIGJ6j9FObXDwKC-lBhUJaxx88XRUNXYaVi-yiQuitN4g/exec";
fetch(`${url}?query=2348`).then(r => r.text()).then(t => console.log(t.slice(0, 500))).catch(e => console.error(e));
