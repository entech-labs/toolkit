# Adding a tool here

This repository is public because clients read it before running what is in it. That makes the
bar for publishing higher than for internal work, not lower.

## Before anything is added

1. **It must be read-only.** No tool here writes to, changes or deletes anything in a client
   environment. If a tool needs to change something, it does not belong in this repository.
2. **It must transmit nothing.** Output is written locally. The client decides what leaves.
3. **It must store no credential.** Interactive sign-in only.
4. **Every safety claim must be checkable.** Do not write "it makes no changes" — write which
   string to search for and what the reader will find. If a claim cannot be verified from the
   source, remove the claim rather than soften it.
5. **No client is named.** Not in code, not in comments, not in sample data, not in a commit
   message. Client-specific framing lives in the covering email, never here. Sample data uses
   obviously fictional names.
6. **It is tested before it is published.** Including the failure paths — a tool that silently
   skips a section is worse than one that fails loudly.

## Review

Two people read it before it goes public: one for what the tool does, one specifically checking
points 1 to 5 above against the source rather than against the README.

## Releases

Tag a version, attach the file, publish the SHA256 in the release notes. Clients are told to
check the hash, so it has to be there.
