# Mock ActiveSync server

A small server for testing what the add-on does when the server changes its
folders, which a real account does not let a script do. It answers OPTIONS,
FolderSync and Sync (every collection empty) and accepts any credentials.
It uses the add-on's own WBXML codec, so no packages are needed.

```
node mock/server.mjs [port]        # default 47660
```

Add an account in TbSync: type "custom", server `http://127.0.0.1:47660`,
any user name and password. Then change the server from a script:

```
curl -X POST localhost:47660/__mock/rename -d '{"serverId":"mycal","name":"Team events"}'
curl -X POST localhost:47660/__mock/rename -d '{"serverId":"mycal","parentId":"trash"}'   # move into the trash
curl -X POST localhost:47660/__mock/add    -d '{"name":"Clients","type":14}'
curl -X POST localhost:47660/__mock/delete -d '{"serverId":"mybook"}'
curl localhost:47660/__mock/folders        # the tree
curl localhost:47660/__mock/log            # every request seen, decoded
curl -X POST localhost:47660/__mock/reset  # back to the start
```

The next sync of the account picks the change up as a FolderSync Update, Add
or Delete. The starting tree has the three default folders, a trash folder and
one user-created folder of each kind (`mycal`, `mybook`, `mytasks`); folder
types are the EAS ones (7 tasks, 8 calendar, 9 contacts, 13/14/15 the
user-created kinds, 4 deleted items).
