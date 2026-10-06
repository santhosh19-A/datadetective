const path = require('path');
const cp = require('child_process');

const serverPath = path.join(__dirname, 'server.js');

// server.js is the single source of truth. This launcher no longer
// rewrites the backend or applies regex patches at startup.
const child = cp.spawn(process.execPath, [serverPath], {
    stdio: 'inherit',
});

child.on('exit', (code, signal) => {
    if (signal) {
        console.log(`server.js exited from signal ${signal}`);
    } else {
        console.log(`server.js exited with code ${code}`);
    }
});
