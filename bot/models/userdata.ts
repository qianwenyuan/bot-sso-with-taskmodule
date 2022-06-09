class Userdata {
    data;

    constructor() {
    }

    input(data) {
        this.data = data;
    }

    getme() {
        return this.data;
    }
};

export const userdata = new Userdata();