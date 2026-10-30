# The Node version tracks `.node-version`, which is what CI's primary leg runs. Renovate
# bumps this tag; move the major only together with `.node-version` and `engines`.
FROM node:26.13.0-trixie

# What CI's render-oracle job installs, plus `fonts-liberation`: the ubuntu-24.04 runner
# image ships it, and the render smoke relies on Arial resolving to Liberation Sans.
# `sudo` is for `playwright install --with-deps`, which installs Chromium's system libraries.
RUN apt-get update \
	&& apt-get install -y --no-install-recommends libreoffice-impress poppler-utils fonts-liberation sudo \
	&& rm -rf /var/lib/apt/lists/* \
	&& echo 'node ALL=(root) NOPASSWD:ALL' > /etc/sudoers.d/node \
	&& chmod 0440 /etc/sudoers.d/node

# The GitHub CLI from GitHub's own apt repository, not Debian's: `ooxml-validate` verifies its
# download with `gh attestation verify`, which Debian's older build does not have.
RUN curl -fsSL https://cli.github.com/packages/githubcli-archive-keyring.gpg -o /usr/share/keyrings/githubcli-archive-keyring.gpg \
	&& echo 'deb [signed-by=/usr/share/keyrings/githubcli-archive-keyring.gpg] https://cli.github.com/packages stable main' > /etc/apt/sources.list.d/github-cli.list \
	&& apt-get update \
	&& apt-get install -y --no-install-recommends gh \
	&& rm -rf /var/lib/apt/lists/*

# pnpm comes from the `packageManager` field in package.json, through corepack.
ENV COREPACK_ENABLE_DOWNLOAD_PROMPT=0
RUN corepack enable
