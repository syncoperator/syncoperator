const Storage = {
    KEY: 'QS_PRO_V5',
    
    save(data) {
        localStorage.setItem(this.KEY, JSON.stringify(data));
    },
    
    load() {
        const data = localStorage.getItem(this.KEY);
        return data ? JSON.parse(data) : {
            fields: { lzf: '', sag: '', stt: '', stn: '', abs: '', grf: '' },
            tools: []
        };
    }
};
