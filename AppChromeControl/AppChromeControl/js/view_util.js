'use strict';
function DOMUtil() {
        DOMUtil.prototype.div = function (classlist) {
            var d = document.createElement('div');
            if (classlist === undefined) {
                return d;
            }
            if (typeof classlist === 'string') {
                addClasses(d, [classlist]);
            } else {
                addClasses(d, classlist);
            }
            return d;
        };

        DOMUtil.prototype.create = function (type, classlist) {
            var d = document.createElement(type);
            if (classlist === undefined) {
                return d;
            }
            if (typeof classlist === 'string') {
                addClasses(d, [classlist]);
            } else {
                addClasses(d, classlist);
            }
            return d;
        };

        DOMUtil.prototype.attr = function (element, attribute, value) {
            try {
                if (typeof attribute === 'string' && typeof value === 'string') {
                    element.setAttribute(attribute, value);
                } else {
                    for (var key in attribute) {
                        if (attribute.hasOwnProperty(key)) {
                            element.setAttribute(key, attribute[key]);
                        }
                    }
                }
            } catch (e) {
                throw (e);
            }
        };

        DOMUtil.prototype.mount = function (father, child) {
            try {
                if (father != null && child != null) {
                    father.appendChild(child);
                }
            } catch (e) {
                throw (e);
            }
        };

        DOMUtil.prototype.unmount = function (father, child) {
            try {
                if (father != null && child != null) {
                    father.removeChild(child);
                }
            } catch (e) {
                throw (e);
            }
        };

        function addClasses(element, classes) {
            try {
                var found = false;
                if (classes instanceof Array && element instanceof HTMLElement) {
                    classes.forEach(
                    function (className) {
                        var classes = element.className.split(' ');
                        var j = classes.length;
                        while (j--) {
                            if (classes[j] === className) {
                                found = true;
                            }
                        }
                        if (!found) {
                            classes.push(className);
                        }
                        element.className = classes.join(' ');
                    });
                } else {
                    console.warn("element or classes type mismatch");
                }
            } catch (e) {
                throw (e);
            }
        }
        
        DOMUtil.prototype.classes = addClasses;
        
        DOMUtil.prototype.containsClasses = function (element, classes) {
            var found =  false;
            try {
                if (classes instanceof Array && element instanceof HTMLElement) {
                    classes.forEach(
                    function (className) {
                        var classes = element.className.split(' ');
                        var j = classes.length;
                        while (j--) {
                            if (classes[j] === className) {
                                found = true;
                            }
                        }
                    });
                } else {
                    console.warn("element or classes type mismatch");
                }
            } catch (e) {
                throw (e);
            }
            return found;
        };

        DOMUtil.prototype.removeClasses = function (element, classes) {
            try {
                if (classes instanceof Array && element instanceof HTMLElement) {
                    classes.forEach(
                    function (className) {
                        var classes = element.className.split(' ');
                        var j = classes.length;
                        while (j--) {
                            if (classes[j] === className) {
                                classes.splice(j, 1);
                            }
                        }
                        element.className = classes.join(' ');
                    });
                } else {
                    console.warn("element or classes type mismatch");
                }
            } catch (e) {
                throw (e);
            }
        };

        var createElem = function (define) {
            var type = define.type;
            if (type == null || type.length == 0) {
                return null;
            }
            var elem = document.createElement(type);
            var attr = define.attr;
            if (attr) {
                for (var key in attr) {
                    if (attr.hasOwnProperty(key)) {
                        elem.setAttribute(key, attr[key]);
                    }
                }
            }
            var content = define.content;
            if (content && content.length != 0) {
                elem.textContent = content;
            }
            var subElem = define.subElem;
            if (subElem && subElem.length != 0) {
                for (var key in subElem) {
                    if (subElem[key]) {
                        var em = createElem(subElem[key]);
                        if (em) {
                            elem.appendChild(em);
                        }
                    }
                }
            }
            var action = define.action;
            if (action) {
                elem.onclick = action;
            }
            var keyAction = define.keyAction;
            if (keyAction) {
                elem.onkeydown = keyAction;
            }
            if (define.oninput) {
                elem.oninput = define.oninput;
            }
            if (define.onchange) {
                elem.onchange = define.onchange;
            }
            if (define.onkeypress) {
                elem.onkeypress = elem.onkeypress;
            }
            if (define.onpaste) {
                elem.onpaste = define.onpaste;
            }
            if (define.onerror) {
                elem.onerror = define.onerror;
            }
            return elem;
        };
        DOMUtil.prototype.createElem = createElem;
    }

module.exports = DOMUtil;