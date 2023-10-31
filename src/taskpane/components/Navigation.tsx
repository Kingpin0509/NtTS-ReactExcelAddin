import * as React from 'react';
import { Nav, INavLink, INavStyles, INavLinkGroup } from '@fluentui/react/lib/Nav';
import { DefaultButton } from '@fluentui/react/lib/Button';

// Define the nav items
const navLinkGroups: INavLinkGroup[] = [
  {
    name: 'Home',
    expandAriaLabel: 'Expand Home section',
    collapseAriaLabel: 'Collapse Home section',
    links: [
      {
        name: 'Activity',
        url: 'http://example.com',
        key: 'key1',
        target: '_blank',
      },
      {
        name: 'News',
        url: 'http://msn.com',
        key: 'key2',
        target: '_blank',
      },
    ],
  },
  {
    name: 'Documents',
    expandAriaLabel: 'Expand Documents section',
    collapseAriaLabel: 'Collapse Documents section',
    links: [
      {
        name: 'Word',
        url: 'http://example.com',
        key: 'key3',
        target: '_blank',
      },
      {
        name: 'Excel',
        url: 'http://example.com',
        key: 'key4',
        target: '_blank',
      },
    ],
  },
];

// Define a custom component to render the links as buttons
const LinkAsButton = (props: INavLink) => {
    return (
      <DefaultButton
        text={props.name}
        href={props.url}
        target={props.target}
        onClick={props.onClick}
        checked={props.isSelected}
      />
    );
  };

// Define the nav props
const navProps = {
    // The nav items
    groups: navLinkGroups,
  
    // The custom component to render the links
    linkAs: LinkAsButton,
  };
  
// Define the nav styles
const navStyles: Partial<INavStyles> = {
  root: {
    width: 208,
    height: 350,
    boxSizing: 'border-box',
    border: '1px solid #eee',
    overflowY: 'auto',
  },
};

// // Define the nav props
// const navProps = {
//   // The nav items
//   groups: navLinkGroups,

//   // The callback function when a link is clicked
//   onLinkClick: (ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => {
//     if (item && item.name === 'News') {
//       alert('News link clicked');
//     }
//   },

//   // The selected key of the nav item
//   selectedKey: 'key3',

//   // The aria label for the nav container
//   ariaLabel: 'Nav example',

//   // The custom styles for the nav component
//   styles: navStyles,
// };

// Create a function component that uses the nav props
export const NavCustomLinkExample = () => {
  return <Nav {...navProps} />;
};