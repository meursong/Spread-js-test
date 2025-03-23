export const useLoading = () => {
    const loading = ref(false)

    const startLoading = () => {
        loading.value = true
    }

    const stopLoading = () => {
        loading.value = false
    }

    return {
        loading: readonly(loading),
        startLoading,
        stopLoading
    }
}
