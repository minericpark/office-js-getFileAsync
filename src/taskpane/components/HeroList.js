import * as React from "react";
import PropTypes from "prop-types";

function HeroList(props) {
  const { children, items, message } = props;

  const listItems = items.map((item, index) => (
    <li className="ms-ListItem" key={index}>
      <i className={`ms-Icon ms-Icon--${item.icon}`} />
      <span className="ms-font-m ms-fontColor-neutralPrimary">{item.primaryText}</span>
    </li>
  ));

  return (
    <main className="ms-welcome__main">
      <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{message}</h2>
      <ul className="ms-List ms-welcome__features ms-u-slideUpIn10">{listItems}</ul>
      {children}
    </main>
  );
}

export default HeroList;

HeroList.propTypes = {
  children: PropTypes.node,
  items: PropTypes.array,
  message: PropTypes.string,
};
