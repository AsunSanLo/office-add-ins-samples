/**
 * Calcular cambio de moneda
 * @customfunction
 * @param value Valor
 * @param currency Moneda
 * @param invocation Custom function handler
 * @returns El valor actual en la moneda
 */
/* global clearInterval, console, setInterval */

export function currency(value: number, currency: string, invocation: CustomFunctions.StreamingInvocation<string>): void {
  invocation.setResult("LOADING...");
  fetch("https://api.exchangeratesapi.io/latest").then(response => response.json())
  .then(data => {
    const result = data.rates[currency] * value;
    invocation.setResult(`${result.toFixed(2)} ${currency}`);
  })
  .catch(()=> {
    invocation.setResult("UPS, error");
  });
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {

  invocation.setResult("LOADING...");
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}
