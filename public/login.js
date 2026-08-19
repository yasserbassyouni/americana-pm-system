document.addEventListener(
    "DOMContentLoaded",
    () => {

        // ====================================================
        // ELEMENTS
        // ====================================================

        const loginForm =
            document.getElementById(
                "loginForm"
            );


        const usernameInput =
            document.getElementById(
                "username"
            );


        const passwordInput =
            document.getElementById(
                "password"
            );


        const loginButton =
            document.getElementById(
                "loginButton"
            );


        const loginButtonText =
            document.getElementById(
                "loginButtonText"
            );


        const loginLoading =
            document.getElementById(
                "loginLoading"
            );


        const loginMessage =
            document.getElementById(
                "loginMessage"
            );


        const togglePassword =
            document.getElementById(
                "togglePassword"
            );


        // ====================================================
        // CHECK EXISTING SESSION
        // ====================================================

        checkExistingSession();


        // ====================================================
        // PASSWORD SHOW / HIDE
        // ====================================================

        togglePassword.addEventListener(
            "click",
            () => {

                if (
                    passwordInput.type ===
                    "password"
                ) {

                    passwordInput.type =
                        "text";

                    togglePassword.textContent =
                        "Hide";

                    togglePassword.title =
                        "Hide password";

                } else {

                    passwordInput.type =
                        "password";

                    togglePassword.textContent =
                        "Show";

                    togglePassword.title =
                        "Show password";
                }

            }
        );


        // ====================================================
        // LOGIN
        // ====================================================

        loginForm.addEventListener(
            "submit",
            async (
                event
            ) => {

                event.preventDefault();


                clearMessage();


                const username =
                    usernameInput
                        .value
                        .trim();


                const password =
                    passwordInput
                        .value;


                if (
                    !username ||
                    !password
                ) {

                    showMessage(
                        "Please enter your username and password.",
                        "error"
                    );

                    return;
                }


                setLoading(
                    true
                );


                try {

                    const response =
                        await fetch(
                            "/api/auth/login",
                            {
                                method:
                                    "POST",

                                headers: {
                                    "Content-Type":
                                        "application/json"
                                },

                                credentials:
                                    "include",

                                body:
                                    JSON.stringify(
                                        {
                                            username,
                                            password
                                        }
                                    )
                            }
                        );


                    const data =
                        await response.json();


                    if (
                        !response.ok
                    ) {

                        throw new Error(
                            data.error ||
                            data.message ||
                            "Login failed."
                        );
                    }


                    showMessage(
                        "Login successful. Opening PM System...",
                        "success"
                    );


                    setTimeout(
                        () => {

                            window.location.href =
                                "/";

                        },
                        500
                    );

                } catch (
                    error
                ) {

                    showMessage(
                        error.message ||
                        "Unable to sign in.",
                        "error"
                    );

                    setLoading(
                        false
                    );
                }

            }
        );


        // ====================================================
        // CHECK SESSION
        // ====================================================

        async function checkExistingSession() {

            try {

                const response =
                    await fetch(
                        "/api/auth/me",
                        {
                            credentials:
                                "include"
                        }
                    );


                if (
                    response.ok
                ) {

                    const data =
                        await response.json();


                    if (
                        data.authenticated ===
                        true
                    ) {

                        window.location.href =
                            "/";
                    }
                }

            } catch (
                error
            ) {

                // User is not logged in.
                // Stay on login page.

            }

        }


        // ====================================================
        // LOADING STATE
        // ====================================================

        function setLoading(
            loading
        ) {

            loginButton.disabled =
                loading;


            usernameInput.disabled =
                loading;


            passwordInput.disabled =
                loading;


            if (
                loading
            ) {

                loginButtonText.classList.add(
                    "hidden"
                );

                loginLoading.classList.remove(
                    "hidden"
                );

            } else {

                loginButtonText.classList.remove(
                    "hidden"
                );

                loginLoading.classList.add(
                    "hidden"
                );
            }

        }


        // ====================================================
        // MESSAGE
        // ====================================================

        function showMessage(
            message,
            type
        ) {

            loginMessage.textContent =
                message;


            loginMessage.className =
                "login-message";


            loginMessage.classList.add(
                type
            );


            loginMessage.classList.remove(
                "hidden"
            );

        }


        function clearMessage() {

            loginMessage.textContent =
                "";


            loginMessage.className =
                "login-message hidden";

        }


        // ====================================================
        // FOCUS USERNAME
        // ====================================================

        usernameInput.focus();

    }
);