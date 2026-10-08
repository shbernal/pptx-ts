# PowerPoint VM

A Windows 11 VM with desktop PowerPoint, run under [dockur/windows](https://github.com/dockur/windows)
on a Linux host. It serves the PowerPoint worker (`scripts/powerpoint/worker.mjs`) on
`127.0.0.1:8765`, so PowerPoint can act as a remote oracle while development stays on Linux.
On Windows or macOS, use your own PowerPoint instead.

Everything except the Office sign-in is automatic. The container runs only while you need it.

## Requirements

- Linux with KVM (`/dev/kvm`) and `/dev/net/tun`, and Docker with Compose.
- About 64 GB of free disk for the image, and the RAM and cores set in `.env` (6 GB and 4 cores by
  default).
- A Microsoft 365 account that includes the desktop apps, and a Windows license. dockur installs
  with Microsoft's generic trial key. Licensing is yours to get right.
- The Office product in `oem/office.xml` must match that account. It installs
  `O365ProPlusRetail` (Microsoft 365 Apps for enterprise), which comes with E3/E5 and the
  education A3/A5 plans. A Personal or Family account needs `O365HomePremRetail`, and Business
  Standard needs `O365BusinessRetail`. Change it before the first boot.

## Bring-up

1. Copy `.env.example` to `.env` and set `VM_PASSWORD`.
2. Run `pnpm ppt:vm:up`. It generates the worker token into `.env`, stages the worker in
   `shared/`, creates `storage/` with copy-on-write off (on btrfs a VM image under
   copy-on-write fragments and can break Windows Setup), and starts the container.
3. Open <http://127.0.0.1:8006> and watch. The first run downloads Windows and installs it
   unattended. That takes a while. At the first logon, `oem/install.bat` runs in a console
   window. It installs Node 24 and PowerShell 7, opens the worker's port in the guest firewall,
   turns off the lock screen and screen saver, sets the Office no-update policy, installs
   Microsoft 365 Apps with the Office Deployment Tool (`oem/install-office.cmd`, about 3 GB),
   and registers the worker's logon task. Its log is `C:\OEM\install.log`. If the Office step
   failed, run `C:\OEM\install-office.cmd` again.
4. **Manual:** open PowerPoint once in the VM, sign in with the Microsoft 365 account so it
   activates, and dismiss the first-run dialogs. A dialog left open blocks every COM call after
   it.
5. Reboot the VM with `pnpm ppt:vm:sync`, which restarts a running container. Windows signs in on
   its own and the logon task starts the worker.
6. Run `pnpm ppt:health`. It prints the PowerPoint build the worker drives.
7. Back the VM up, as described below.

### Office updates

`install.bat` sets the Office "Enable Automatic Updates" policy to off before Office is
installed. Fixtures record the PowerPoint build that authored them, and an unattended update
would change it in the middle of a series. To move to a newer build on purpose, use
**File > Account > Update Options > Update Now** in PowerPoint, then rerun `pnpm ppt:health`.

## Day to day

| Command | What it does |
|---|---|
| `pnpm ppt:vm:up` | Stages the worker and starts the VM |
| `pnpm ppt:vm:down` | Shuts Windows down cleanly and removes the container. The disk stays in `storage/` |
| `pnpm ppt:vm:sync` | Re-stages the worker from the repo and restarts the VM if it is running |
| `pnpm ppt:health` | Reports the worker's PowerPoint build, or why it cannot be reached |
| `pnpm run test:com` | Runs the PowerPoint COM smoke through the worker |

`.env` sets `TSPPTX_POWERPOINT_URL`, which sends `test:com` to the worker instead of looking for
a local PowerPoint. The environment variable of the same name overrides it.

### Re-syncing the worker

After changing anything under `scripts/powerpoint/`, run `pnpm ppt:vm:sync`. Windows sees
`shared/` as `\\host.lan\Data`. Each time `oem/start-worker.cmd` starts the worker, it mirrors
`\\host.lan\Data\ts-pptx-worker` to `C:\ts-pptx-worker`. The restart is how the new files get
loaded. Add `--no-restart` to stage the files without restarting.

### Connecting to the desktop

Use the web viewer at <http://127.0.0.1:8006>. RDP also works, on `127.0.0.1:3389` as
`VM_USERNAME`. Disconnecting an RDP session locks the console session, though, and PowerPoint
automation needs an unlocked desktop. After using RDP, reboot with `pnpm ppt:vm:sync`, or sign
back in through the web viewer.

## Backup

The whole VM is the disk image in `storage/`. Stop the VM so the image is consistent, then copy
the directory:

```sh
pnpm ppt:vm:down
cp -r --sparse=always --preserve=timestamps tools/powerpoint-vm/storage ~/backups/powerpoint-vm-$(date +%F)
```

The container owns `storage/` as root, but its files are world-readable, so the copy needs no
`sudo`. The image is a 64 GB sparse file. A provisioned VM occupies about 16 GB, and
`--sparse=always` keeps the copy that size.

To restore, put the copy back as `storage/` and run `pnpm ppt:vm:up`.

## Security

The worker executes whatever script a request sends it. The bearer token is the only thing
between a caller and arbitrary code execution as the VM's user. `compose.yml` publishes every
port on `127.0.0.1` only. Keep it that way, and never forward these ports to another machine.

## Files

| File | Purpose |
|---|---|
| `compose.yml` | The container: Windows edition, resources, local-only ports, volumes |
| `.env.example` | The settings `.env` takes. `.env` is gitignored and holds the token |
| `oem/install.bat` | One-time provisioning at the first logon |
| `oem/install-office.cmd` | Installs Microsoft 365 Apps with the Office Deployment Tool; rerunnable |
| `oem/office.xml` | The Office Deployment Tool configuration: product, apps, no updates |
| `oem/start-worker.cmd` | The logon task: mirrors the worker from the share and keeps it running |
| `storage/` | The disk image (gitignored) |
| `shared/` | Staged worker files and token (gitignored) |
