## Setup on a new Windows machine
1. Install VSCode, Node, Git, Git LFS.
2. Install the VC++ runtime (needed by Jest's native resolver):
   `winget install Microsoft.VCRedist.2015+.x64`
3. `npm ci`