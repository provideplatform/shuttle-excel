/* eslint-disable @typescript-eslint/no-unused-vars */
import * as $ from "jquery";
import { encodeForHTML } from "./validate";

var listArray = [];
// State
// Number of products
var numberOfItems;
var numberPerPage;
var currentPage;

// Number of pages
var numberOfPages;

export const paginate = (items: any[], listName: string) => {
  console.log('paginate items', items)

  listArray = [];
  var parentElement = document.getElementById(`${listName}`);
  parentElement.innerHTML = "";

  items.forEach((item, index) => {
    var newButton = document.createElement("button");
    newButton.setAttribute("type", "button");
    newButton.setAttribute("class", "list-group-item list-group-item-action");
    newButton.setAttribute("style", "display: none");
    // listName === "mappings-list"
    //   ? newButton.setAttribute("id", encodeForHTML(item))
    //   : newButton.setAttribute("id", encodeForHTML(item.id));
    newButton.setAttribute("id", encodeForHTML(item.id))
    newButton.textContent = item.name;
    listArray.push(newButton);
    parentElement.appendChild(newButton);
  });

  // State
  // Number of products
  numberOfItems = listArray.length;
  numberPerPage = 5;
  currentPage = 1;

  // Number of pages
  numberOfPages = Math.ceil(numberOfItems / numberPerPage);

  buildPage(1, listName);
  buildPagination(currentPage, listName);

  $(`#${listName}-paginator`).on("click", "button", function () {
    var clickedPage = parseInt($(this).val().toString());
    buildPagination(clickedPage, listName);
    buildPage(clickedPage, listName);
  });
};

function accomodatePage(clickedPage) {
  if (clickedPage <= 1) {
    return clickedPage + 1;
  }
  if (clickedPage >= numberOfPages) {
    return clickedPage - 1;
  }
  return clickedPage;
}

function buildPagination(clickedPage, listName) {
  var paginator = document.getElementById(`${listName}-paginator`);
  var innerHTMLContent = "";
  paginator.innerHTML = "";

  const currPageNum = accomodatePage(clickedPage);
  if (numberOfPages >= 3) {
    for (let i = -1; i < 2; i++) {
      innerHTMLContent += `<button class="btn bg-transparent m-1 text-light" value="${currPageNum + i}">${
        currPageNum + i
      }</button>`;
    }
    paginator.innerHTML = innerHTMLContent;
  } else {
    for (let i = 0; i < numberOfPages; i++) {
      innerHTMLContent += `<button class="btn bg-transparent m-1 text-light" value="${i + 1}">${i + 1}</button>`;
    }
    paginator.innerHTML = innerHTMLContent;
  }
}

function buildPage(currPage, listName) {
  const trimStart = (currPage - 1) * numberPerPage;
  const trimEnd = trimStart + numberPerPage;

  var paginatedList = document.getElementById(`${listName}`).children;

  for (var i = 0; i < paginatedList.length; i++) {
    if (trimStart <= i && i < trimEnd) {
      paginatedList[i].setAttribute("style", "");
    } else {
      paginatedList[i].setAttribute("style", "display: none");
    }
  }
}
