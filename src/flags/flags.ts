interface Country {
    name: string;
    code: string;
  }
  
  const countries: Country[] = [
    { name: "Afghanistan", code: "af" },
    { name: "Albania", code: "al" },
    { name: "Algeria", code: "dz" },
    { name: "Andorra", code: "ad" },
    { name: "Angola", code: "ao" },
    { name: "Antigua and Barbuda", code: "ag" },
    { name: "Argentina", code: "ar" },
    { name: "Armenia", code: "am" },
    { name: "Australia", code: "au" },
    { name: "Austria", code: "at" },
    { name: "Azerbaijan", code: "az" },
    { name: "Bahamas", code: "bs" },
    { name: "Bahrain", code: "bh" },
    { name: "Bangladesh", code: "bd" },
    { name: "Barbados", code: "bb" },
    { name: "Belarus", code: "by" },
    { name: "Belgium", code: "be" },
    { name: "Belize", code: "bz" },
    { name: "Benin", code: "bj" },
    { name: "Bhutan", code: "bt" },
    { name: "Bolivia", code: "bo" },
    { name: "Bosnia and Herzegovina", code: "ba" },
    { name: "Botswana", code: "bw" },
    { name: "Brazil", code: "br" },
    { name: "Brunei", code: "bn" },
    { name: "Bulgaria", code: "bg" },
    { name: "Burkina Faso", code: "bf" },
    { name: "Burundi", code: "bi" },
    { name: "Cabo Verde", code: "cv" },
    { name: "Cambodia", code: "kh" },
    { name: "Cameroon", code: "cm" },
    { name: "Canada", code: "ca" },
    { name: "Central African Republic", code: "cf" },
    { name: "Chad", code: "td" },
    { name: "Chile", code: "cl" },
    { name: "China", code: "cn" },
    { name: "Colombia", code: "co" },
    { name: "Comoros", code: "km" },
    { name: "Congo (Congo-Brazzaville)", code: "cg" },
    { name: "Costa Rica", code: "cr" },
    { name: "Croatia", code: "hr" },
    { name: "Cuba", code: "cu" },
    { name: "Cyprus", code: "cy" },
    { name: "Czechia (Czech Republic)", code: "cz" },
    { name: "Democratic Republic of the Congo", code: "cd" },
    { name: "Denmark", code: "dk" },
    { name: "Djibouti", code: "dj" },
    { name: "Dominica", code: "dm" },
    { name: "Dominican Republic", code: "do" },
    { name: "Ecuador", code: "ec" },
    { name: "Egypt", code: "eg" },
    { name: "El Salvador", code: "sv" },
    { name: "Equatorial Guinea", code: "gq" },
    { name: "Eritrea", code: "er" },
    { name: "Estonia", code: "ee" },
    { name: "Eswatini", code: "sz" },
    { name: "Ethiopia", code: "et" },
    { name: "Fiji", code: "fj" },
    { name: "Finland", code: "fi" },
    { name: "France", code: "fr" },
    { name: "Gabon", code: "ga" },
    { name: "Gambia", code: "gm" },
    { name: "Georgia", code: "ge" },
    { name: "Germany", code: "de" },
    { name: "Ghana", code: "gh" },
    { name: "Greece", code: "gr" },
    { name: "Grenada", code: "gd" },
    { name: "Guatemala", code: "gt" },
    { name: "Guinea", code: "gn" },
    { name: "Guinea-Bissau", code: "gw" },
    { name: "Guyana", code: "gy" },
    { name: "Haiti", code: "ht" },
    { name: "Holy See", code: "va" },
    { name: "Honduras", code: "hn" },
    { name: "Hungary", code: "hu" },
    { name: "Iceland", code: "is" },
    { name: "India", code: "in" },
    { name: "Indonesia", code: "id" },
    { name: "Iran", code: "ir" },
    { name: "Iraq", code: "iq" },
    { name: "Ireland", code: "ie" },
    { name: "Israel", code: "il" },
    { name: "Italy", code: "it" },
    { name: "Jamaica", code: "jm" },
    { name: "Japan", code: "jp" },
    { name: "Jordan", code: "jo" },
    { name: "Kazakhstan", code: "kz" },
    { name: "Kenya", code: "ke" },
    { name: "Kiribati", code: "ki" },
    { name: "Kuwait", code: "kw" },
    { name: "Kyrgyzstan", code: "kg" },
    { name: "Laos", code: "la" },
    { name: "Latvia", code: "lv" },
    { name: "Lebanon", code: "lb" },
    { name: "Lesotho", code: "ls" },
    { name: "Liberia", code: "lr" },
    { name: "Libya", code: "ly" },
    { name: "Liechtenstein", code: "li" },
    { name: "Lithuania", code: "lt" },
    { name: "Luxembourg", code: "lu" },
    { name: "Madagascar", code: "mg" },
    { name: "Malawi", code: "mw" },
    { name: "Malaysia", code: "my" },
    { name: "Maldives", code: "mv" },
    { name: "Mali", code: "ml" },
    { name: "Malta", code: "mt" },
    { name: "Marshall Islands", code: "mh" },
    { name: "Mauritania", code: "mr" },
    { name: "Mauritius", code: "mu" },
    { name: "Mexico", code: "mx" },
    { name: "Micronesia", code: "fm" },
    { name: "Moldova", code: "md" },
    { name: "Monaco", code: "mc" },
    { name: "Mongolia", code: "mn" },
    { name: "Montenegro", code: "me" },
    { name: "Morocco", code: "ma" },
    { name: "Mozambique", code: "mz" },
    { name: "Myanmar", code: "mm" },
    { name: "Namibia", code: "na" },
    { name: "Nauru", code: "nr" },
    { name: "Nepal", code: "np" },
    { name: "Netherlands", code: "nl" },
    { name: "New Zealand", code: "nz" },
    { name: "Nicaragua", code: "ni" },
    { name: "Niger", code: "ne" },
    { name: "Nigeria", code: "ng" },
    { name: "North Korea", code: "kp" },
    { name: "North Macedonia", code: "mk" },
    { name: "Norway", code: "no" },
    { name: "Oman", code: "om" },
    { name: "Pakistan", code: "pk" },
    { name: "Palau", code: "pw" },
    { name: "Palestine", code: "ps" },
    { name: "Panama", code: "pa" },
    { name: "Papua New Guinea", code: "pg" },
    { name: "Paraguay", code: "py" },
    { name: "Peru", code: "pe" },
    { name: "Philippines", code: "ph" },
    { name: "Poland", code: "pl" },
    { name: "Portugal", code: "pt" },
    { name: "Qatar", code: "qa" },
    { name: "Romania", code: "ro" },
    { name: "Russia", code: "ru" },
    { name: "Rwanda", code: "rw" },
    { name: "Saint Kitts and Nevis", code: "kn" },
    { name: "Saint Lucia", code: "lc" },
    { name: "Saint Vincent and the Grenadines", code: "vc" },
    { name: "Samoa", code: "ws" },
    { name: "San Marino", code: "sm" },
    { name: "Sao Tome and Principe", code: "st" },
    { name: "Saudi Arabia", code: "sa" },
    { name: "Senegal", code: "sn" },
    { name: "Serbia", code: "rs" },
    { name: "Seychelles", code: "sc" },
    { name: "Sierra Leone", code: "sl" },
    { name: "Singapore", code: "sg" },
    { name: "Slovakia", code: "sk" },
    { name: "Slovenia", code: "si" },
    { name: "Solomon Islands", code: "sb" },
    { name: "Somalia", code: "so" },
    { name: "South Africa", code: "za" },
    { name: "South Korea", code: "kr" },
    { name: "South Sudan", code: "ss" },
    { name: "Spain", code: "es" },
    { name: "Sri Lanka", code: "lk" },
    { name: "Sudan", code: "sd" },
    { name: "Suriname", code: "sr" },
    { name: "Sweden", code: "se" },
    { name: "Switzerland", code: "ch" },
    { name: "Syria", code: "sy" },
    { name: "Tajikistan", code: "tj" },
    { name: "Tanzania", code: "tz" },
    { name: "Thailand", code: "th" },
    { name: "Timor-Leste", code: "tl" },
    { name: "Togo", code: "tg" },
    { name: "Tonga", code: "to" },
    { name: "Trinidad and Tobago", code: "tt" },
    { name: "Tunisia", code: "tn" },
    { name: "Turkey", code: "tr" },
    { name: "Turkmenistan", code: "tm" },
    { name: "Tuvalu", code: "tv" },
    { name: "Uganda", code: "ug" },
    { name: "Ukraine", code: "ua" },
    { name: "United Arab Emirates", code: "ae" },
    { name: "United Kingdom", code: "gb" },
    { name: "United States of America", code: "us" },
    { name: "Uruguay", code: "uy" },
    { name: "Uzbekistan", code: "uz" },
    { name: "Vanuatu", code: "vu" },
    { name: "Venezuela", code: "ve" },
    { name: "Vietnam", code: "vn" },
    { name: "Yemen", code: "ye" },
    { name: "Zambia", code: "zm" },
    { name: "Zimbabwe", code: "zw" }
  ];
  
  function initialize() {
    console.log("Initializing Flag Generator...");
    populateDropdown();
    setupDropdownChange();
  }
  
  function populateDropdown() {
    const selectElem = document.getElementById("countrySelect") as HTMLSelectElement;
    if (!selectElem) {
      console.error("Dropdown element 'countrySelect' not found.");
      return;
    }
    // Clear any existing options.
    selectElem.innerHTML = "";
    // Add a default option.
    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = "Select a country";
    selectElem.appendChild(defaultOption);
    // Add an option for each country.
    countries.forEach((country) => {
      const option = document.createElement("option");
      option.value = country.code;
      option.textContent = country.name;
      selectElem.appendChild(option);
    });
    console.log("Dropdown populated with", countries.length, "countries.");
  }
  
  function setupDropdownChange() {
    const selectElem = document.getElementById("countrySelect") as HTMLSelectElement;
    if (!selectElem) {
      console.error("Dropdown element 'countrySelect' not found.");
      return;
    }
    selectElem.addEventListener("change", () => {
      const selectedCode = selectElem.value;
      if (!selectedCode) {
        hideFlagPreview();
        return;
      }
      // Construct the flag URL using FlagCDN with a width of 320.
      const flagUrl = `https://flagcdn.com/w1280/${selectedCode.toLowerCase()}.png`;
      console.log("Selected country code:", selectedCode, "Flag URL:", flagUrl);
      displayFlag(selectElem.options[selectElem.selectedIndex].text, flagUrl);
    });
  }
  
  function displayFlag(countryName: string, flagUrl: string) {
    const flagSection = document.getElementById("flagSection");
    const flagTitle = document.getElementById("flagTitle");
    const flagImage = document.getElementById("flagImage") as HTMLImageElement;
    if (flagSection && flagTitle && flagImage) {
      flagTitle.textContent = countryName;
      flagImage.src = flagUrl;
      flagImage.alt = `${countryName} Flag`;
      flagSection.classList.remove("hidden");
    } else {
      console.error("Flag preview elements not found.");
    }
  }
  
  function hideFlagPreview() {
    const flagSection = document.getElementById("flagSection");
    if (flagSection) {
      flagSection.classList.add("hidden");
    }
  }
  
  // Use DOMContentLoaded as a fallback, then Office.onReady.
  document.addEventListener("DOMContentLoaded", () => {
    console.log("DOM fully loaded. Initializing...");
    initialize();
  });
  
  Office.onReady(() => {
    console.log("Office is ready. Initializing...");
    initialize();
  });
  