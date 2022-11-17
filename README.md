# Life-Work Calendar-Sync

This Bookmarklet-App downloads your `Office365 Outlook` / `Microsoft Teams` calendar as iCalender `*.ics` document.
You can import those events into your own calender to schedule your `Life-Work-Balance`.

A `Bookmarklet-App` is some kind of robot executed in the context of a website in your browser.

```html
<a href="javascript: (() => { alert(document.cookie); })();">
  Get Biscuit.
</a>
```

## Development

 * Get the source code `git clone https://gist.github.com/evo42/a5c0d9b6a18b15e0387b48430bcdd46b ./iCalSync && cd ./iCalSync`
 * Run the setup and build scripts `npm run setup` and `npm run build`
 * Change the source code.
 * Run `npm run build` again.