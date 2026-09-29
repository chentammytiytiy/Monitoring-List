const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyKAKdT8uMsLjhKEXYJueCuN72uggHihcBSy396BHirI6bQqz592fB-v4Y0DAmgzyfILQ/exec';
async function run() {
    const r = await fetch(SCRIPT_URL);
    const text = await r.text();
    console.log("Returned size:", text.length);
    console.log("Snippet:", text.substring(0, 500));
}
run();
