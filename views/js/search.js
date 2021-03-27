var timer;
var inputQuery = $('#inputQuery');
inputQuery.on('keyup input', function () {
  timer && clearTimeout(timer);
  timer = setTimeout(updatePage, 5000);
});
$('#entitySearch').submit(function (event) {
  updatePage();
  event.preventDefault();
});

function updatePage() {
  var searchVal = inputQuery.val().trim();
  var currentUrl = window.location.href;
  var newUrl = replaceUrlParam(currentUrl, 'q', searchVal);
  window.location.href = newUrl;
}

// http://stackoverflow.com/questions/7171099/how-to-replace-url-parameter-with-javascript-jquery#20420424
function replaceUrlParam(url, paramName, paramValue) {
  var pattern = new RegExp('\\b(' + paramName + '=).*?(&|$)');
  if (url.search(pattern) >= 0) {
    return url.replace(pattern, '$1' + paramValue + '$2');
  }
  return (
    url +
    (url.indexOf('?') > 0 ? '&' : '?') +
    paramName +
    '=' +
    encodeURIComponent(paramValue)
  );
}
