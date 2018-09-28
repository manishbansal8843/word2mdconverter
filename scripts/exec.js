const cp = require("child_process");
const isWin = /^win/.test(process.platform);


module.exports = exec;

function exec(cmd, args, options = {}) {
    return /**@type {Promise<{exitCode:number}>}*/(new Promise((resolve, reject) => {
         const subshellFlag = isWin ? "/c" : "-c";
        const command = isWin ? [possiblyQuote(cmd), ...args] : [`${cmd} ${args.join(" ")}`];
        
        const proc = cp.spawn(isWin ? "cmd" : "/bin/sh", [subshellFlag, ...command], { stdio: "inherit", windowsVerbatimArguments: true });
      
        proc.on("exit", exitCode => {
           
            if (exitCode === 0 ) {
                resolve({ exitCode });
            }
            else {
                reject(new Error(`Process exited with code: ${exitCode}`));
            }
        });
        proc.on("error", error => {
            reject(error);
        });
    }));
}

/**
 * @param {string} cmd
 */
function possiblyQuote(cmd) {
    return cmd.indexOf(" ") >= 0 ? `"${cmd}"` : cmd;
}
