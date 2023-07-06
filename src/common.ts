const icons = new Map<string, string>([
    ["App Store", "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' fill='none' viewBox='0 0 24 24'><path d='M 19.666667,0 A 4.3333333,4.3333333 0 0 1 24,4.3333333 V 19.666667 A 4.3333333,4.3333333 0 0 1 19.666667,24 H 4.3333333 A 4.3333333,4.3333333 0 0 1 0,19.666667 V 4.3333333 A 4.3333333,4.3333333 0 0 1 4.3333333,0 Z M 8.492,16.338667 H 6.1733333 L 6.084,16.493333 6.024,16.614667 a 1,1 0 0 0 1.7106667,1 l 0.076,-0.113334 0.68,-1.162666 z m 5.232,-6.9733337 -1.161333,1.9893337 3.544,6.144 0.07467,0.113333 a 1,1 0 0 0 1.717334,-0.990667 l -0.06,-0.122666 -0.669334,-1.16 H 18.336 l 0.136,-0.008 A 1,1 0 0 0 19.326667,14.476 l 0.0093,-0.136 -0.0093,-0.136 A 1,1 0 0 0 18.472,13.349333 L 18.336,13.34 h -2.32 L 13.722667,9.3666667 Z M 13.684,5.468 A 1,1 0 0 0 12.392,5.7146667 L 12.316,5.828 11.996,6.3733333 11.684,5.832 11.609333,5.72 a 1,1 0 0 0 -1.156,-0.32 l -0.134666,0.066667 -0.113334,0.074667 a 1,1 0 0 0 -0.3186663,1.156 L 9.9533337,6.8306667 10.836,8.3613333 7.928,13.34 H 5.6693333 l -0.136,0.0093 a 1,1 0 0 0 0,1.981334 l 0.136,0.0093 h 8.0359997 l -1.153333,-2 -2.308,-0.0013 3.8,-6.502667 0.06,-0.1213333 A 1,1 0 0 0 13.684,5.468Z'></path></svg>"],
    ["Power Apps", "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' fill='none' viewBox='0 0 2048 2048'><path d='M1760 256q60 0 112 22t92 62 61 91 23 113v768q0 55-19 104t-54 87-82 63-101 31v-128q28-6 51-20t41-36 26-46 10-55V544q0-33-12-62t-35-51-51-34-62-13H288q-33 0-62 12t-51 35-34 51-13 62v768q0 29 9 54t27 47 40 35 52 21v128q-54-6-101-30t-81-63-54-88-20-104V544q0-60 22-112t62-92 91-61 113-23h1472zm-253 1332q0 26-19 45l-137 137q-19 19-46 19t-46-19l-137-137q-19-19-19-45 0-27 19-46l137-137q19-19 46-19t46 19l137 137q19 19 19 46zm285-286q0 26-19 45l-137 137q-19 19-46 19-26 0-45-19l-137-137q-19-19-19-45 0-27 19-46l137-138q19-19 45-19 27 0 46 19l137 138q19 19 19 46zM734 1789q-26 0-45-19l-379-378q-9-9-20-19t-20-21-16-23-7-27q0-15 6-27t16-24 21-21 20-19l379-378q19-19 45-19 27 0 46 19l378 378q9 9 20 19t21 21 16 23 7 28q0 15-6 27t-16 23-21 21-21 19l-378 378q-19 19-46 19zm202-487q0-27-19-46l-137-138q-19-19-46-19-26 0-45 19l-138 138q-19 19-19 46t18 45 36 35l103 102q19 19 45 19 27 0 46-19l102-102q18-18 36-35t18-45zm369-85q-27 0-46-19l-137-137q-19-19-19-45 0-27 19-46l137-137q19-19 46-19t46 19l137 137q19 19 19 46 0 26-19 45l-137 137q-19 19-46 19z'></path></svg>"],
    ["Power Automate", "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' fill='none' viewBox='0 0 2048 2048'><path d='M1760 256q60 0 112 22t92 62 61 91 23 113v768q0 60-22 112t-62 92-91 61-113 23h-672v-128h672q33 0 62-12t51-35 34-51 13-62V544q0-33-12-62t-35-51-51-34-62-13H288q-33 0-62 12t-51 35-34 51-13 62v672H0V544q0-60 22-112t62-92 91-61 113-23h1472zm-626 896q-19 0-32-13t-14-33V960H973l-225 620q-10 28-27 45t-38 26-45 12-50 3q-20 0-39-1t-37-1v96q0 40-28 68t-68 28H96q-40 0-68-28t-28-68v-320q0-40 28-68t68-28h320q40 0 68 28t28 68v96h115l226-620q10-28 26-45t38-26 45-12 49-3q20 0 39 1t38 1V686q0-19 13-32t33-14h420q19 0 32 13t14 33v420q0 19-13 32t-33 14h-420zm-750 320H128v256h256v-256z'></path></svg>"],
    ["Power BI", "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' fill='none' viewBox='0 0 2048 2048'><path d='M1760 256q60 0 112 22t92 62 61 91 23 113v768q0 56-20 105t-54 87-81 62-101 31v-128q27-6 50-20t41-35 27-47 10-55V544q0-33-12-62t-35-51-51-34-62-13H288q-33 0-62 12t-51 35-34 51-13 62v768q0 33 12 62t35 51 51 34 62 13h32v128h-32q-60 0-112-22t-92-62-61-91-23-113V544q0-60 22-112t62-92 91-61 113-23h1472zM576 1312q27 0 50 10t40 27 28 41 10 50v192q0 27-10 50t-27 40-41 28-50 10q-27 0-50-10t-40-27-28-41-10-50v-192q0-27 10-50t27-40 41-28 50-10zm320-384q27 0 50 10t40 27 28 41 10 50v576q0 27-10 50t-27 40-41 28-50 10q-27 0-50-10t-40-27-28-41-10-50v-576q0-27 10-50t27-40 41-28 50-10zm320 128q27 0 50 10t40 27 28 41 10 50v448q0 27-10 50t-27 40-41 28-50 10q-27 0-50-10t-40-27-28-41-10-50v-448q0-27 10-50t27-40 41-28 50-10zm320-384q27 0 50 10t40 27 28 41 10 50v832q0 27-10 50t-27 40-41 28-50 10q-27 0-50-10t-40-27-28-41-10-50V800q0-27 10-50t27-40 41-28 50-10z'></path></svg>"],
    ["PowerShell", "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' fill='none' viewBox='0 0 2048 2048'><path d='M0 128h2048v1664H0V128zm1920 128H128v128h1792V256zM128 1664h1792V512H128v1152zm768-128v-128h640v128H896zM549 716l521 372-521 372-74-104 375-268-375-268 74-104z'></path></svg>"],
    ["SharePoint", "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' fill='none' viewBox='0 0 2048 2048'><path d='M2048 1152q0 106-40 199t-110 162-163 110-199 41q-65 0-130-17-9 85-47 158t-99 128-137 84-163 31q-93 0-174-35t-142-96-96-142-36-175q0-16 1-32t4-32H85q-35 0-60-25t-25-60V597q0-35 25-60t60-25h302q12-109 62-202t127-163T751 39 960 0q119 0 224 45t183 124 123 183 46 224q0 16-1 32t-3 32q106 0 200 40t164 109 111 162 41 201zM960 128q-83 0-158 29t-135 80-98 122-52 153h422q35 0 60 25t25 60v422l18-3q18-64 52-121t80-104 104-81 122-52q8-43 8-82 0-93-35-174t-96-142-142-96-175-36zM522 1357q46 0 90-9t80-32 56-58 22-90q0-54-22-90t-57-60-73-39-73-29-56-29-23-38q0-17 12-27t28-17 35-8 30-2q51 0 90 13t81 39V729q-24-7-44-12t-41-8-41-5-48-2q-44 0-90 10t-83 32-61 58-24 88q0 51 22 85t57 58 73 40 73 31 56 31 23 39q0 19-10 30t-27 17-33 7-32 2q-60 0-106-20t-95-53v160q102 40 211 40zm438 563q66 0 124-25t101-68 69-102 26-125q0-57-19-109t-53-93-81-71-103-41v165q0 35-25 60t-60 25H646q-6 32-6 64 0 66 25 124t68 102 102 69 125 25zm576-384q79 0 149-30t122-82 83-122 30-150q0-79-30-149t-82-122-123-83-149-30q-79 0-148 30t-122 83-83 122-31 149v22q0 11 2 22 47 23 87 55t71 73 54 88 33 98q67 26 137 26z'></path></svg>"],
    ["Teams", "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' fill='none' viewBox='0 0 2048 2048'><path d='M1760 704q-47 0-87-17t-71-48-48-71-18-88q0-46 17-87t48-71 71-48 88-18q46 0 87 17t71 48 48 72 18 87q0 47-17 87t-48 71-72 48-87 18zm0-320q-40 0-68 28t-28 68q0 40 28 68t68 28q40 0 68-28t28-68q0-40-28-68t-68-28zm288 480v476q0 66-25 124t-68 102-102 69-125 25q-38 0-77-9t-73-28q-25 81-73 147t-112 114-143 74-162 26q-98 0-184-34t-154-94-112-142-58-178H85q-35 0-60-25t-25-60V597q0-35 25-60t60-25h733q-29-61-29-128 0-62 23-116t64-95 95-64 117-24q62 0 116 23t95 64 64 95 24 117q0 62-23 116t-64 95-95 64-117 24q-16 0-32-2t-32-5v92h928q40 0 68 28t28 68zm-960-651q-35 0-66 13t-55 37-36 55-14 66q0 35 13 66t37 55 54 36 67 14q35 0 66-13t54-37 37-54 14-67q0-35-13-66t-37-54-55-37-66-14zM592 848h192V688H240v160h192v512h160V848zm880 624V896h-448v555q0 35-25 60t-60 25H709q13 69 47 128t84 101 113 67 135 24q79 0 149-30t122-82 83-122 30-150zm448-132V896h-320v585q26 26 59 38t69 13q40 0 75-15t61-41 41-61 15-75z'></path></svg>"]
]);

// Returns an icon as an SVG element
export function getIcon(height?, width?, iconName?, className?) {
    // Get the icon element
    let elDiv = document.createElement("div");
    elDiv.innerHTML = iconName ? icons.get(iconName) : icons.get('App Store');
    let icon = elDiv.firstChild as SVGImageElement;
    if (icon) {
        // See if a class name exists
        if (className) {
            // Parse the class names
            let classNames = className.split(' ');
            for (let i = 0; i < classNames.length; i++) {
                // Add the class name
                icon.classList.add(classNames[i]);
            }
        } else {
            icon.classList.add("icon-svg");
        }
        // Set the height/width
        height ? icon.setAttribute("height", (height).toString()) : null;
        width ? icon.setAttribute("width", (width).toString()) : null;
        // Update the styling
        icon.style.pointerEvents = "none";
        // Support for IE
        icon.setAttribute("focusable", "false");
    }
    // Return the icon
    return icon;
}