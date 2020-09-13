class DepartmentsWP {
  constructor() {
    this.url = `${_spPageContextInfo.webServerRelativeUrl}/`;
    this.departmentsBody = 'dprtBody';
    this.overlay = document.getElementById('overlay-register');
    this.friendSelect = document.querySelector('[name="favFriend"]');
  }

  async getUsers() {
    let result;
    try {
      const query = `${this.url}_api/web/lists/getbytitle('ansapMainList')/items?$select=Title,isActive,Gender,favFriendCharacterId`;
      result = await this.getItems(query);
    } catch (err) {
      console.log(err);
    }
    return result.d.results;
  }

  async getFriends() {
    let result;
    try {
      const query = `${this.url}_api/web/lists/getbytitle('friendsList')/items?$select=Title,Id`;
      result = await this.getItems(query);
    } catch (err) {
      console.log(err);
    }
    return result.d.results;
  }

  async getItems(query) {
    return $.ajax({
      url: query,
      method: 'GET',
      contentType: 'application/json;odata=verbose',
      headers: {
        Accept: 'application/json;odata=verbose',
      },
    });
  }

  async createUser(webUrl, newUser) {
    const query = `${webUrl}_api/web/lists/getbytitle('ansapMainList')/items`;
    const requestDigest = await this.getRequestDigest(webUrl);
    const listItemType = await this.getListItemType(webUrl, 'ansapMainList');
    const objType = {
      __metadata: {
        type: listItemType.d.ListItemEntityTypeFullName,
      },
    };
    const objData = JSON.stringify(Object.assign(objType, newUser));

    return $.ajax({
      url: query,
      type: 'POST',
      data: objData,
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest':
          requestDigest.d.GetContextWebInformation.FormDigestValue,
        'X-HTTP-Method': 'POST',
      },
    });
  }

  getRequestDigest(webUrl) {
    return $.ajax({
      url: webUrl + '_api/contextinfo',
      method: 'POST',
      headers: {
        Accept: 'application/json; odata=verbose',
      },
    });
  }

  getListItemType(url, listTitle) {
    const query =
      url +
      "_api/Web/Lists/getbytitle('" +
      listTitle +
      "')/ListItemEntityTypeFullName";
    return this.getItems(query);
  }

  renderHTML(users, friends) {
    try {
      let departmentsItems = '';
      users.map((item) => {
        const departmentItem = new DepartmentItem(item, friends);
        departmentsItems += departmentItem.getHTML();
      });
      document.getElementById(
        this.departmentsBody
      ).innerHTML = departmentsItems;
    } catch (err) {
      console.log(err);
    }

    let friendsOptions =
      '<option selected="selected" disabled>Favorite Friends Character</option>';
    friends.map((item) => {
      friendsOptions += this.createFriendsMarkup(item);
    });
    this.friendSelect.innerHTML = friendsOptions;

    // Sign up modal:
    document.getElementById('register').addEventListener('click', (e) => {
      e.preventDefault();
      this.openModal();
    });

    this.overlay.addEventListener('click', (e) => {
      if (
        e.target === e.currentTarget ||
        e.target.dataset.action === 'close-modal'
      ) {
        this.closeModal();
      }
    });
  }

  openModal() {
    this.overlay.classList.remove('hide-modal');
    this.overlay.classList.add('show-modal');

    document
      .getElementById('signup-btn')
      .addEventListener('click', async () => {
        const username = document.querySelector('[name="usename"]').value;
        const gender = document.querySelector('[name="gender"]').value;
        const favFriend = document.querySelector('[name="favFriend"]').value;
        const newUser = {
          Title: username,
          isActive: true,
          Gender: gender || 'male',
          favFriendCharacterId: Number(favFriend) || 1,
        };
        if (username === '') return;
        try {
          await this.createUser(this.url, newUser);
          this.closeModal();
        } catch (err) {
          console.log(err);
        }
      });
  }

  closeModal() {
    this.overlay.classList.add('hide-modal');
    this.overlay.classList.remove('show-modal');
  }

  createFriendsMarkup(item) {
    return `<option value="${item.Id}">${item.Title}</option>`;
  }
}

class DepartmentItem {
  constructor(departmentItem, friends) {
    this.username = departmentItem.Title;
    this.gender = departmentItem.Gender;
    this.online = departmentItem.isActive;
    this.friendId = departmentItem.favFriendCharacterId;
    this.friendsList = friends;
  }

  getHTML() {
    return `
      <div class="${this.online ? 'dprtCard' : 'dprtCardOffline'}">
        <img class="dprtCard__img" src="../SiteAssets/ansap/avatar.svg" alt="avatar"/>
        <h3 class="dprtCard__name">@${this.username}</h3>
        <p>Gender: ${this.gender}</p>
        <p class="isOnline">${this.online ? 'Online' : 'Offline'}</p>
        <p class="favFriend">
          <span>Favorite Friends Character: </span>
          <span>${
            this.friendsList.find((o) => o.Id === this.friendId).Title
          }</span>
        </p>
      </div>`;
  }
}

SP.SOD.executeFunc('sp.js', 'SP.ClientContext', async function () {
  const dprtWP = new DepartmentsWP();
  const users = await dprtWP.getUsers();
  const friends = await dprtWP.getFriends();
  dprtWP.renderHTML(users, friends);
});
