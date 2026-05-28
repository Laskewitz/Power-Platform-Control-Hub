import { useState } from 'react';
import {
  Toast,
  ToastBody,
  ToastTitle,
  useToastController,
} from '@fluentui/react-components';

interface UseMutationOptions<TResult> {
  successMessage?: string;
  onSuccess?: (result: TResult) => void;
  onError?: (error: Error) => void;
}

export function useMutation<TArgs extends unknown[], TResult = void>(
  fn: (...args: TArgs) => Promise<TResult>,
  options?: UseMutationOptions<TResult>,
): { execute: (...args: TArgs) => Promise<TResult | undefined>; isLoading: boolean } {
  const [isLoading, setIsLoading] = useState(false);
  const { dispatchToast } = useToastController('coe-toaster');

  async function execute(...args: TArgs): Promise<TResult | undefined> {
    setIsLoading(true);

    try {
      const result = await fn(...args);

      if (options?.successMessage) {
        dispatchToast(
          <Toast>
            <ToastTitle>Success</ToastTitle>
            <ToastBody>{options.successMessage}</ToastBody>
          </Toast>,
          { intent: 'success' },
        );
      }

      options?.onSuccess?.(result);
      return result;
    } catch (error: unknown) {
      const resolvedError = error instanceof Error
        ? error
        : new Error('An unexpected error occurred.');

      dispatchToast(
        <Toast>
          <ToastTitle>Action failed</ToastTitle>
          <ToastBody>{resolvedError.message}</ToastBody>
        </Toast>,
        { intent: 'error' },
      );

      options?.onError?.(resolvedError);
      return undefined;
    } finally {
      setIsLoading(false);
    }
  }

  return { execute, isLoading };
}
